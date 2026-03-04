using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace Tabify
{
    public partial class MainWindow : Window
    {
        // =====================================================================
        //  データモデル
        // =====================================================================
        private class TabData
        {
            public IntPtr HWnd;
            public string Title;
            public Border Header;
            public TextBlock TitleText;
            public TranslateTransform Shift = new();
        }

        private readonly List<TabData> _tabs = new();
        private int _selectedIndex = -1;

        // タブ設定
        private double _tabW = MaxTabW;
        private const double MaxTabW = 200.0;
        private const double MinTabW = 48.0;
        private const double TabGap = 2.0;
        private const double TabH = 36.0;
        private const double CaptionBtnSpace = 150.0;

        // D&D 状態
        private TabData _dragTab;
        private int _dragInsertIndex;
        private double _dragOffsetX;
        private bool _isDraggingTab;
        private Point _dragStartPoint;
        private const double DragThreshold = 5.0;

        // ウィンドウ移動・リサイズ中フラグ
        private bool _isWindowMoving;

        // Win32
        private NativeMethods.WinEventDelegate _winEventProc;
        private IntPtr _hookID = IntPtr.Zero;
        private IntPtr _capturedHWnd = IntPtr.Zero;
        private readonly Dictionary<IntPtr, IntPtr> _origStyles = new();
        private readonly Dictionary<IntPtr, IntPtr> _origParents = new();

        private bool IsForeground =>
            new WindowInteropHelper(this).Handle == GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private readonly DispatcherTimer _syncTimer;
        private readonly DispatcherTimer _highlightTimer;

        // =====================================================================
        //  初期化
        // =====================================================================
        public MainWindow()
        {
            InitializeComponent();

            _syncTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) };
            _syncTimer.Tick += (_, _) => { if (!_isWindowMoving) SyncActiveWindow(); };
            _syncTimer.Start();

            _highlightTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(50) };
            _highlightTimer.Tick += HighlightTimer_Tick;

            _winEventProc = WinEventProc;
            _hookID = NativeMethods.SetWinEventHook(
                NativeMethods.EVENT_SYSTEM_MOVESIZESTART,
                NativeMethods.EVENT_SYSTEM_MOVESIZEEND,
                IntPtr.Zero, _winEventProc, 0, 0,
                NativeMethods.WINEVENT_OUTOFCONTEXT);

            Activated += (_, _) => BringSelectedToFront();
            StateChanged += (_, _) => UpdateMaximizeIcon();
        }

        // WndProc フック
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            if (PresentationSource.FromVisual(this) is HwndSource src)
                src.AddHook(WndProcHook);
        }

        private IntPtr WndProcHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            switch (msg)
            {
                case NativeMethods.WM_ENTERSIZEMOVE:
                    _isWindowMoving = true;
                    break;
                case NativeMethods.WM_EXITSIZEMOVE:
                    _isWindowMoving = false;
                    ApplySync(DispatcherPriority.Render);
                    break;
                case NativeMethods.WM_MOVE:
                case NativeMethods.WM_SIZE:
                    if (_isWindowMoving) ApplySync(DispatcherPriority.Render);
                    break;
            }
            return IntPtr.Zero;
        }

        private void ApplySync(DispatcherPriority priority)
        {
            Dispatcher.BeginInvoke(priority, new Action(SyncActiveWindow));
        }

        // =====================================================================
        //  ハイライト演出
        // =====================================================================
        private void HighlightTimer_Tick(object sender, EventArgs e)
        {
            if (_capturedHWnd == IntPtr.Zero)
            {
                _highlightTimer.Stop();
                UpdateHighlight(0);
                return;
            }

            if (NativeMethods.GetCursorPos(out var pt))
            {
                try
                {
                    var local = TabBarBg.PointFromScreen(new Point(pt.X, pt.Y));
                    bool isOver = local.X >= 0 && local.X <= TabBarBg.ActualWidth &&
                                  local.Y >= 0 && local.Y <= TabBarBg.ActualHeight;
                    UpdateHighlight(isOver ? 0.6 : 0);
                }
                catch { UpdateHighlight(0); }
            }
        }

        private void UpdateHighlight(double opacity)
        {
            if (Math.Abs(DropHighlight.Opacity - opacity) < 0.01) return;
            var anim = new DoubleAnimation(opacity, TimeSpan.FromMilliseconds(150));
            DropHighlight.BeginAnimation(OpacityProperty, anim);
        }

        // =====================================================================
        //  外部ウィンドウ検知
        // =====================================================================
        private void WinEventProc(IntPtr hook, uint evt, IntPtr hwnd,
            int idObj, int idChild, uint thread, uint time)
        {
            var me = new WindowInteropHelper(this).Handle;
            if (hwnd == me) return;
            NativeMethods.GetWindowThreadProcessId(hwnd, out uint pid);
            if (pid == (uint)System.Diagnostics.Process.GetCurrentProcess().Id) return;

            if (evt == NativeMethods.EVENT_SYSTEM_MOVESIZESTART)
            {
                _capturedHWnd = hwnd;
                _highlightTimer.Start();
            }
            else if (evt == NativeMethods.EVENT_SYSTEM_MOVESIZEEND && _capturedHWnd == hwnd)
            {
                if (NativeMethods.GetCursorPos(out var pt))
                {
                    try
                    {
                        var local = TabBarBg.PointFromScreen(new Point(pt.X, pt.Y));
                        if (local.X >= 0 && local.X <= TabBarBg.ActualWidth &&
                            local.Y >= 0 && local.Y <= TabBarBg.ActualHeight)
                            AttachWindow(hwnd);
                    }
                    catch { }
                }
                _capturedHWnd = IntPtr.Zero;
                _highlightTimer.Stop();
                UpdateHighlight(0);
            }
        }

        private void AttachWindow(IntPtr hWnd)
        {
            foreach (var t in _tabs) if (t.HWnd == hWnd) return;

            if (IsUwpWindow(hWnd))
            {
                ShowDropHintTemporary("⚠ Microsoft Store アプリは統合できません");
                return;
            }

            if (!_origStyles.ContainsKey(hWnd))
            {
                var style = NativeMethods.GetWindowLongPtr(hWnd, NativeMethods.GWL_STYLE);
                _origStyles[hWnd] = style;
                var ns = style.ToInt64() & ~(NativeMethods.WS_THICKFRAME | NativeMethods.WS_CAPTION);
                NativeMethods.SetWindowLongPtr(hWnd, NativeMethods.GWL_STYLE, new IntPtr(ns));

                var op = NativeMethods.GetWindowLongPtr(hWnd, NativeMethods.GWLP_HWNDPARENT);
                _origParents[hWnd] = op;
                NativeMethods.SetWindowLongPtr(hWnd, NativeMethods.GWLP_HWNDPARENT,
                    new WindowInteropHelper(this).Handle);
                
                NativeMethods.SetWindowPos(hWnd, IntPtr.Zero, 0, 0, 0, 0,
                    NativeMethods.SWP_NOZORDER | NativeMethods.SWP_FRAMECHANGED);
                NativeMethods.ShowWindow(hWnd, NativeMethods.SW_RESTORE);
            }

            var tab = new TabData { HWnd = hWnd, Title = GetTitle(hWnd) };
            tab.Header = BuildTabHeader(tab);
            tab.Header.RenderTransform = tab.Shift;

            foreach (var t in _tabs) NativeMethods.ShowWindow(t.HWnd, NativeMethods.SW_HIDE);
            
            _tabs.Add(tab);
            TabBar.Children.Add(tab.Header);
            DropHintText.Visibility = Visibility.Collapsed;

            RecalcTabWidths(animate: false);
            SelectTab(_tabs.Count - 1);
            LayoutTabs(animate: false);

            SyncActiveWindow();
        }

        // =====================================================================
        //  タブヘッダー生成
        // =====================================================================
        private Border BuildTabHeader(TabData tab)
        {
            tab.TitleText = new TextBlock
            {
                Text = tab.Title,
                TextTrimming = TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = Brushes.DarkGray,
                FontSize = 12,
                MaxWidth = _tabW - 36
            };

            var closeText = new TextBlock
            {
                Text = "✕", FontSize = 9,
                Foreground = new SolidColorBrush(Color.FromRgb(0xaa, 0xaa, 0xaa)),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            var closeBtn = new Border
            {
                Width = 20, Height = 20, CornerRadius = new CornerRadius(10),
                Background = Brushes.Transparent,
                Margin = new Thickness(5, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Cursor = Cursors.Arrow,
                Child = closeText
            };
            closeBtn.MouseEnter += (_, _) =>
                closeBtn.Background = new SolidColorBrush(Color.FromArgb(0x60, 0xff, 0x55, 0x55));
            closeBtn.MouseLeave += (_, _) => closeBtn.Background = Brushes.Transparent;
            closeBtn.MouseLeftButtonDown += (s, e) => e.Handled = true;
            closeBtn.MouseLeftButtonUp += (s, e) => { e.Handled = true; CloseApp(tab); };

            var row = new Grid();
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetColumn(tab.TitleText, 0);
            Grid.SetColumn(closeBtn, 1);
            row.Children.Add(tab.TitleText);
            row.Children.Add(closeBtn);

            var border = new Border
            {
                Width = _tabW, Height = TabH,
                CornerRadius = new CornerRadius(8, 8, 0, 0),
                Background = new SolidColorBrush(Color.FromRgb(0x3c, 0x3c, 0x3c)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0x28, 0x28, 0x28)),
                BorderThickness = new Thickness(1, 1, 1, 0),
                Padding = new Thickness(12, 0, 8, 0),
                Cursor = Cursors.Hand,
                Tag = tab, Child = row
            };
            border.MouseEnter += (_, _) =>
            {
                if (_tabs.IndexOf(tab) != _selectedIndex)
                    border.Background = new SolidColorBrush(Color.FromRgb(0x4a, 0x4a, 0x4a));
            };
            border.MouseLeave += (_, _) =>
            {
                if (_tabs.IndexOf(tab) != _selectedIndex)
                    border.Background = new SolidColorBrush(Color.FromRgb(0x3c, 0x3c, 0x3c));
            };
            border.MouseLeftButtonDown += Tab_MouseLeftButtonDown;

            var ctx = new ContextMenu();
            var detach = new MenuItem { Header = "このタブを切り離す (Detach)" };
            detach.Click += (_, _) => DetachTab(tab);
            ctx.Items.Add(detach);
            border.ContextMenu = ctx;

            return border;
        }

        // =====================================================================
        //  タブ選択
        // =====================================================================
        private void SelectTab(int index)
        {
            if (index < 0 || index >= _tabs.Count) return;
            
            for (int i = 0; i < _tabs.Count; i++)
            {
                bool sel = i == index;
                var t = _tabs[i];
                t.Header.Background = new SolidColorBrush(
                    sel ? Color.FromRgb(0x1e, 0x1e, 0x1e) : Color.FromRgb(0x3c, 0x3c, 0x3c));
                t.Header.Margin = new Thickness(0, sel ? 2 : 6, 0, 0);
                if (t.TitleText != null)
                    t.TitleText.Foreground = new SolidColorBrush(
                        sel ? Colors.White : Color.FromRgb(0xaa, 0xaa, 0xaa));

                if (!sel) NativeMethods.ShowWindow(t.HWnd, NativeMethods.SW_HIDE);
            }

            _selectedIndex = index;
            SyncActiveWindow();
        }

        // =====================================================================
        //  タブ幅・レイアウト
        // =====================================================================
        private void RecalcTabWidths(bool animate = true)
        {
            if (_tabs.Count == 0 || TabBar.ActualWidth < 1) { _tabW = MaxTabW; return; }
            double space = TabBar.ActualWidth - 10;
            double targetW = Math.Max(MinTabW, Math.Min(MaxTabW, space / _tabs.Count));
            if (Math.Abs(targetW - _tabW) < 0.5) return;
            _tabW = targetW;

            foreach (var t in _tabs)
            {
                if (animate)
                    t.Header.BeginAnimation(WidthProperty, new DoubleAnimation(targetW, TimeSpan.FromMilliseconds(150)));
                else
                    t.Header.Width = targetW;
                
                if (t.TitleText != null) t.TitleText.MaxWidth = Math.Max(10, targetW - 40);
            }
        }

        private void TabBar_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            RecalcTabWidths();
            LayoutTabs(animate: false);
        }

        private void LayoutTabs(bool animate = true)
        {
            var dur = TimeSpan.FromMilliseconds(150);
            var ease = new CubicEase { EasingMode = EasingMode.EaseOut };
            for (int i = 0; i < _tabs.Count; i++)
            {
                double x = i * (_tabW + TabGap) + 2;
                Canvas.SetTop(_tabs[i].Header, 2);
                if (animate)
                    _tabs[i].Header.BeginAnimation(Canvas.LeftProperty,
                        new DoubleAnimation { To = x, Duration = dur, EasingFunction = ease });
                else
                    Canvas.SetLeft(_tabs[i].Header, x);
            }
        }

        // =====================================================================
        //  タブ D&D 
        // =====================================================================
        private void Tab_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border b && b.Tag is TabData tab)
            {
                _dragTab = tab;
                _dragStartPoint = e.GetPosition(TabBar);
                _dragOffsetX = e.GetPosition(b).X;
                _dragInsertIndex = _tabs.IndexOf(tab);
                _isDraggingTab = false;
                b.CaptureMouse();
                e.Handled = true;
            }
        }

        private void TabBar_MouseMove(object sender, MouseEventArgs e)
        {
            if (_dragTab == null || e.LeftButton != MouseButtonState.Pressed) return;
            var pos = e.GetPosition(TabBar);

            if (!_isDraggingTab)
            {
                if (Math.Abs(pos.X - _dragStartPoint.X) < DragThreshold) return;
                _isDraggingTab = true;
                SelectTab(_tabs.IndexOf(_dragTab));
                Canvas.SetZIndex(_dragTab.Header, 999);
            }

            double minX = 2, maxX = 2 + (_tabs.Count - 1) * (_tabW + TabGap);
            double newLeft = Math.Max(minX, Math.Min(maxX, pos.X - _dragOffsetX));
            Canvas.SetLeft(_dragTab.Header, newLeft);

            int newIns = (int)((newLeft - 2 + _tabW / 2) / (_tabW + TabGap));
            newIns = Math.Max(0, Math.Min(_tabs.Count - 1, newIns));
            if (newIns != _dragInsertIndex) { _dragInsertIndex = newIns; AnimateSlots(); }
        }

        private void TabBar_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_dragTab == null) return;
            _dragTab.Header.ReleaseMouseCapture();

            if (_isDraggingTab)
            {
                int from = _tabs.IndexOf(_dragTab), to = _dragInsertIndex;
                if (from != to) { _tabs.RemoveAt(from); _tabs.Insert(to, _dragTab); _selectedIndex = to; }
                foreach (var t in _tabs)
                {
                    t.Shift.BeginAnimation(TranslateTransform.XProperty, null);
                    t.Shift.X = 0;
                    Canvas.SetZIndex(t.Header, 0);
                }
                LayoutTabs(animate: true);
            }
            else
                SelectTab(_tabs.IndexOf(_dragTab));

            _dragTab = null;
            _isDraggingTab = false;
            e.Handled = true;
        }

        private void AnimateSlots()
        {
            int dragFrom = _tabs.IndexOf(_dragTab);
            for (int i = 0; i < _tabs.Count; i++)
            {
                if (_tabs[i] == _dragTab) continue;
                double shift = 0;
                if (dragFrom < _dragInsertIndex && i > dragFrom && i <= _dragInsertIndex)
                    shift = -(_tabW + TabGap);
                else if (dragFrom > _dragInsertIndex && i >= _dragInsertIndex && i < dragFrom)
                    shift = _tabW + TabGap;
                _tabs[i].Shift.BeginAnimation(TranslateTransform.XProperty,
                    new DoubleAnimation { To = shift, Duration = TimeSpan.FromMilliseconds(120) });
            }
        }

        private void TabBarBg_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _isWindowMoving = true;
            try { DragMove(); }
            finally { _isWindowMoving = false; SyncActiveWindow(); }
        }

        // =====================================================================
        //  キャプションボタン
        // =====================================================================
        private void CaptionBtn_Down(object sender, MouseButtonEventArgs e) => e.Handled = true;
        private void BtnMinimize_Click(object sender, MouseButtonEventArgs e) { WindowState = WindowState.Minimized; e.Handled = true; }
        private void BtnMaximize_Click(object sender, MouseButtonEventArgs e) { 
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized; 
            e.Handled = true;
        }
        private void BtnClose_Click(object sender, MouseButtonEventArgs e) { Close(); e.Handled = true; }
        private void UpdateMaximizeIcon() { MaximizeIcon.Text = WindowState == WindowState.Maximized ? "❐" : "□"; }

        private void CaptionBtn_Enter(object sender, MouseEventArgs e) => ((Border)sender).Background = new SolidColorBrush(Color.FromArgb(0x30, 0xff, 0xff, 0xff));
        private void CloseBtn_Enter(object sender, MouseEventArgs e) => ((Border)sender).Background = new SolidColorBrush(Color.FromArgb(0xe0, 0xe8, 0x11, 0x23));
        private void CaptionBtn_Leave(object sender, MouseEventArgs e) => ((Border)sender).Background = Brushes.Transparent;

        // =====================================================================
        //  操作 
        // =====================================================================
        private void CloseApp(TabData tab)
        {
            var hWnd = tab.HWnd;
            _tabs.Remove(tab);
            TabBar.Children.Remove(tab.Header);
            RestoreStyles(hWnd);
            NativeMethods.PostMessage(hWnd, NativeMethods.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            if (_tabs.Count > 0) SelectTab(0); else DropHintText.Visibility = Visibility.Visible;
            RecalcTabWidths();
            LayoutTabs();
        }

        private void DetachTab(TabData tab)
        {
            var hWnd = tab.HWnd;
            _tabs.Remove(tab);
            TabBar.Children.Remove(tab.Header);
            RestoreStyles(hWnd);
            NativeMethods.ShowWindow(hWnd, NativeMethods.SW_RESTORE);
            NativeMethods.GetCursorPos(out var pt);
            NativeMethods.SetWindowPos(hWnd, IntPtr.Zero, pt.X - 100, pt.Y - 10, 800, 600, NativeMethods.SWP_SHOWWINDOW | NativeMethods.SWP_FRAMECHANGED);
            if (_tabs.Count > 0) SelectTab(0); else DropHintText.Visibility = Visibility.Visible;
            RecalcTabWidths();
            LayoutTabs();
        }

        private void RestoreStyles(IntPtr hWnd)
        {
            bool changed = false;
            if (_origStyles.TryGetValue(hWnd, out var s)) { NativeMethods.SetWindowLongPtr(hWnd, NativeMethods.GWL_STYLE, s); _origStyles.Remove(hWnd); changed = true; }
            if (_origParents.TryGetValue(hWnd, out var p)) { NativeMethods.SetWindowLongPtr(hWnd, NativeMethods.GWLP_HWNDPARENT, p); _origParents.Remove(hWnd); changed = true; }
            if (changed) NativeMethods.SetWindowPos(hWnd, IntPtr.Zero, 0, 0, 0, 0, NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOZORDER | NativeMethods.SWP_FRAMECHANGED | NativeMethods.SWP_NOACTIVATE);
        }

        private void BringSelectedToFront()
        {
            if (_selectedIndex >= 0 && _selectedIndex < _tabs.Count)
                NativeMethods.SetWindowPos(_tabs[_selectedIndex].HWnd, IntPtr.Zero, 0, 0, 0, 0, NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);
        }

        private void SyncActiveWindow()
        {
            if (_selectedIndex < 0 || _selectedIndex >= _tabs.Count || !IsLoaded) return;
            var hWnd = _tabs[_selectedIndex].HWnd;
            var helper = new WindowInteropHelper(this);
            if (helper.Handle == IntPtr.Zero) return;
            try
            {
                var pt = ContentArea.PointToScreen(new Point(0, 0));
                var src = PresentationSource.FromVisual(this);
                if (src?.CompositionTarget == null) return;
                double dx = src.CompositionTarget.TransformToDevice.M11;
                double dy = src.CompositionTarget.TransformToDevice.M22;
                int cx = Math.Max(1, (int)(ContentArea.ActualWidth * dx));
                int cy = Math.Max(1, (int)(ContentArea.ActualHeight * dy));
                uint flags = NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_SHOWWINDOW;
                if (!IsForeground) flags |= NativeMethods.SWP_NOZORDER;
                NativeMethods.SetWindowPos(hWnd, IntPtr.Zero, (int)pt.X, (int)pt.Y, cx, cy, flags);
            }
            catch { }
        }

        private void Window_LocationChanged(object sender, EventArgs e) { if (!_isWindowMoving) SyncActiveWindow(); }
        private void Window_SizeChanged(object sender, SizeChangedEventArgs e) { SyncActiveWindow(); }
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_hookID != IntPtr.Zero) NativeMethods.UnhookWinEvent(_hookID);
            _syncTimer.Stop();
            _highlightTimer.Stop();
            foreach (var t in _tabs) { RestoreStyles(t.HWnd); NativeMethods.ShowWindow(t.HWnd, NativeMethods.SW_RESTORE); }
        }

        private static string GetTitle(IntPtr hWnd)
        {
            int len = NativeMethods.GetWindowTextLength(hWnd);
            var sb = new StringBuilder(len + 1);
            NativeMethods.GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private static bool IsUwpWindow(IntPtr hWnd)
        {
            var sb = new StringBuilder(256);
            NativeMethods.GetClassName(hWnd, sb, sb.Capacity);
            return sb.ToString() == "ApplicationFrameWindow";
        }

        private async void ShowDropHintTemporary(string msg)
        {
            var old = DropHintText.Text; DropHintText.Text = msg; DropHintText.Visibility = Visibility.Visible;
            await System.Threading.Tasks.Task.Delay(3000);
            DropHintText.Text = old; DropHintText.Visibility = _tabs.Count > 0 ? Visibility.Collapsed : Visibility.Visible;
        }
    }
}
