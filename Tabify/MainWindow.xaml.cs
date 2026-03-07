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
            public ImageSource Icon;       // ⭐ アプリのアイコン
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

        // タブスクロール
        private double _tabScrollOffset = 0;   // Canvas 内の表示オフセット（px）
        private bool _isScrollMode = false;    // スクロールモード中かどうか

        // プレースメント状態
        private string _tabPlacement = "上側 (Top)";
        private bool IsVertical => _tabPlacement == "左側 (Left)" || _tabPlacement == "右側 (Right)";

        // D&D 状態
        private TabData _dragTab;
        private int _dragInsertIndex;
        private double _dragOffsetX;
        private bool _isDraggingTab;
        private Point _dragStartPoint;
        private const double DragThreshold = 5.0;
        private Window _dragGhostWindow; // ⭐ エリア外ドラッグ時の追尾ゴースト

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
                    var ptLocalTitleBar = TitleBarBg.PointFromScreen(new Point(pt.X, pt.Y));
                    var ptLocalTabBar = TabBarContainer.PointFromScreen(new Point(pt.X, pt.Y));

                    bool isOver = (ptLocalTitleBar.X >= 0 && ptLocalTitleBar.X <= TitleBarBg.ActualWidth &&
                                   ptLocalTitleBar.Y >= 0 && ptLocalTitleBar.Y <= TitleBarBg.ActualHeight) ||
                                  (ptLocalTabBar.X >= 0 && ptLocalTabBar.X <= TabBarContainer.ActualWidth &&
                                   ptLocalTabBar.Y >= 0 && ptLocalTabBar.Y <= TabBarContainer.ActualHeight);
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
                        var ptLocalTitleBar = TitleBarBg.PointFromScreen(new Point(pt.X, pt.Y));
                        var ptLocalTabBar = TabBarContainer.PointFromScreen(new Point(pt.X, pt.Y));
                        
                        // タイトルバー、またはタブバーエリア内のドロップならアタッチ
                        bool isOverTitleBar = ptLocalTitleBar.X >= 0 && ptLocalTitleBar.X <= TitleBarBg.ActualWidth &&
                                              ptLocalTitleBar.Y >= 0 && ptLocalTitleBar.Y <= TitleBarBg.ActualHeight;
                        
                        bool isOverTabBar = ptLocalTabBar.X >= 0 && ptLocalTabBar.X <= TabBarContainer.ActualWidth &&
                                            ptLocalTabBar.Y >= 0 && ptLocalTabBar.Y <= TabBarContainer.ActualHeight;

                        if (isOverTitleBar || isOverTabBar)
                            AttachWindow(hwnd);
                    }
                    catch { }
                }
                _capturedHWnd = IntPtr.Zero;
                _highlightTimer.Stop();
                UpdateHighlight(0);
            }
        }

        private async void AttachWindow(IntPtr hWnd)
        {
            try
            {
                foreach (var t in _tabs) if (t.HWnd == hWnd) return;

                if (IsUwpWindow(hWnd))
                {
                    ShowDropHintTemporary("⚠ Microsoft Store アプリは統合できません");
                    return;
                }

                // ⭐ [バグ修正] エクスプローラーやシステムウィンドウを引き込むとクラッシュするため拒否
                string className = GetClassName(hWnd);
                if (className == "CabinetWClass" || className == "ExploreWClass" || 
                    className == "Progman" || className == "WorkerW" || className == "Shell_TrayWnd")
                {
                    ShowDropHintTemporary("⚠ システムウィンドウ(Explorer等)は統合できません");
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

                // ⭐ 非同期でアイコンを取得（UIを止めないように）
                var iconSource = await GetAppIconAsync(hWnd);

                var tab = new TabData { HWnd = hWnd, Title = GetTitle(hWnd), Icon = iconSource };
                tab.Header = BuildTabHeader(tab);
                tab.Header.RenderTransform = tab.Shift;

                foreach (var t in _tabs) NativeMethods.ShowWindow(t.HWnd, NativeMethods.SW_HIDE);
                
                _tabs.Add(tab);
                TabBar.Children.Add(tab.Header);
                DropHintText.Visibility = Visibility.Collapsed;
                DropHintTextCenter.Visibility = Visibility.Collapsed;

                RecalcTabWidths(animate: false);
                SelectTab(_tabs.Count - 1);
                LayoutTabs(animate: false);

                SyncActiveWindow();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AttachWindow Error: {ex.Message}");
                ShowDropHintTemporary("⚠ ウィンドウの統合に失敗しました");
            }
        }

        // =====================================================================
        //  タブヘッダー生成（UI刷新）
        // =====================================================================
        private Border BuildTabHeader(TabData tab)
        {
            // アイコン
            var img = new Image
            {
                Source = tab.Icon,
                Width = 24, Height = 24,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            tab.TitleText = new TextBlock
            {
                Text = tab.Title,
                TextTrimming = TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = Brushes.DarkGray,
                FontSize = 12,
                MaxWidth = Math.Max(0, _tabW - 56)
            };

            var closeText = new TextBlock
            {
                Text = "✕", FontSize = 9,
                Foreground = new SolidColorBrush(Color.FromRgb(0xcc, 0xcc, 0xcc)),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 0, -1)
            };
            var closeBtn = new Border
            {
                Width = 20, Height = 20, CornerRadius = new CornerRadius(10),
                Background = Brushes.Transparent,
                Margin = new Thickness(3, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Cursor = Cursors.Arrow,
                Child = closeText
            };
            closeBtn.MouseEnter += (_, _) =>
                closeBtn.Background = new SolidColorBrush(Color.FromArgb(0x44, 0xff, 0xff, 0xff));
            closeBtn.MouseLeave += (_, _) => closeBtn.Background = Brushes.Transparent;
            closeBtn.MouseLeftButtonDown += (s, e) => e.Handled = true;
            closeBtn.MouseLeftButtonUp += (s, e) => { e.Handled = true; CloseApp(tab); };

            // タブの中身（アイコン + タイトル + 閉じる）
            var contentStack = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
            if (tab.Icon != null) contentStack.Children.Add(img);
            contentStack.Children.Add(tab.TitleText);
            contentStack.Children.Add(closeBtn);

            var innerBorder = new Border
            {
                Name = "InnerBorder",
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(8, 0, 4, 0),
                Background = Brushes.Transparent,
                VerticalAlignment = VerticalAlignment.Center,
                Child = contentStack
            };

            // 区切り線（右側または下側）
            var separator = new Border
            {
                Name = "Separator",
                Width = 1, Height = 18,
                Background = new SolidColorBrush(Color.FromRgb(0x4d, 0x4d, 0x4d)),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var rootGrid = new Grid();
            rootGrid.Children.Add(innerBorder);
            rootGrid.Children.Add(separator);

            var border = new Border
            {
                Width = _tabW, Height = TabH,
                Background = Brushes.Transparent,
                Cursor = Cursors.Hand,
                Tag = tab, Child = rootGrid,
                ToolTip = tab.Title // 左右配置時のポップアップ用
            };

            border.MouseEnter += (_, _) =>
            {
                if (_tabs.IndexOf(tab) != _selectedIndex)
                    innerBorder.Background = new SolidColorBrush(Color.FromRgb(0x3e, 0x3e, 0x3e));
            };
            border.MouseLeave += (_, _) =>
            {
                if (_tabs.IndexOf(tab) != _selectedIndex)
                    innerBorder.Background = Brushes.Transparent;
            };
            border.MouseLeftButtonDown += Tab_MouseLeftButtonDown;

            var ctx = new ContextMenu();
            var detach = new MenuItem { Header = "このタブを切り離す (Detach)" };
            detach.Click += (_, _) => DetachTab(tab);
            ctx.Items.Add(detach);
            border.ContextMenu = ctx;

            // アタッチ後すぐにアクティブ用のスタイル設定
            if (_tabs.Count == 0) ApplyActiveStyle(border, true);

            return border;
        }

        // アクティブなタブのUIスタイルを適用（区切り線を消す・色を変える）
        private void ApplyActiveStyle(Border tabHeader, bool isActive)
        {
            if (tabHeader?.Child is Grid grid)
            {
                if (grid.Children[0] is Border inner)
                    inner.Background = new SolidColorBrush(isActive ? Color.FromRgb(0x4a, 0x4a, 0x4a) : Colors.Transparent);
                if (grid.Children[1] is Border sep)
                    sep.Visibility = isActive ? Visibility.Collapsed : Visibility.Visible;
            }
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
                ApplyActiveStyle(t.Header, sel);

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
            if (_tabs.Count == 0 || TabBar.ActualWidth < 1 || TabBar.ActualHeight < 1) { _tabW = MaxTabW; return; }

            // 垂直・水平で可変長（Width/Height）を切り替え
            double availableSpace = IsVertical ? TabBar.ActualHeight : TabBar.ActualWidth;
            double scrollBtnSize = _isScrollMode ? 56 : 0;
            double space = availableSpace - scrollBtnSize - 10;
            double idealSize = space / _tabs.Count;
            
            // 垂直（左右メニュー）時は横幅・縦幅ともに広げてアイコンを目立たせる (例:60x60)
            double targetW = IsVertical ? 60.0 : Math.Max(MinTabW, Math.Min(MaxTabW, idealSize));
            double targetH = IsVertical ? 60.0 : TabH;

            // 垂直レイアウト時は理想高さではなく、全体スペースに収まるか（溢れるか）を判定する
            bool needsScroll = IsVertical ? (targetH * _tabs.Count > space) : (idealSize < MinTabW);

            if (needsScroll && !_isScrollMode)
            {
                _isScrollMode = true;
                TabScrollLeft.Visibility = Visibility.Visible;
                TabScrollRight.Visibility = Visibility.Visible;
                TabScrollLeftIcon.Text = IsVertical ? "▲" : "◀";
                TabScrollRightIcon.Text = IsVertical ? "▼" : "▶";
                
                if (IsVertical)
                    TabBar.Margin = new Thickness(0, 56, 0, 0); // 下にスクロール余白を広げる
                else
                    TabBar.Margin = new Thickness(56, 0, _tabPlacement == "上側 (Top)" ? 180 : 10, 0);
                    
                space = (IsVertical ? TabBar.ActualHeight : TabBar.ActualWidth) - 10;
            }
            else if (!needsScroll && _isScrollMode)
            {
                _isScrollMode = false;
                _tabScrollOffset = 0;
                TabScrollLeft.Visibility = Visibility.Collapsed;
                TabScrollRight.Visibility = Visibility.Collapsed;
                TabBar.Margin = new Thickness(0, 0, _tabPlacement == "上側 (Top)" && !IsVertical ? 180 : 0, 0);
            }

            if (Math.Abs(targetW - _tabW) >= 0.5) _tabW = targetW;

            foreach (var t in _tabs)
            {
                if (animate)
                    t.Header.BeginAnimation(WidthProperty, new DoubleAnimation(_tabW, TimeSpan.FromMilliseconds(IsVertical ? 0 : 150)));
                else
                    t.Header.Width = _tabW;

                // テキスト・閉じるボタンの表示制御
                t.TitleText.Visibility = IsVertical ? Visibility.Collapsed : Visibility.Visible;
                if (t.Header.Child is Grid grid && grid.Children.Count > 0 && grid.Children[0] is Border inner && inner.Child is StackPanel stack)
                {
                    if (stack.Children.Count >= 3)
                    {
                        var iconImg = stack.Children[0] as Image;
                        if (iconImg != null)
                        {
                            iconImg.Margin = IsVertical ? new Thickness(0) : new Thickness(0, 0, 6, 0);
                            iconImg.Width = IsVertical ? 24 : 16;
                            iconImg.Height = IsVertical ? 24 : 16;
                        }
                        
                        var closeBtn = stack.Children[stack.Children.Count - 1];
                        closeBtn.Visibility = IsVertical ? Visibility.Collapsed : Visibility.Visible;
                    }
                    inner.Padding = IsVertical ? new Thickness(0) : new Thickness(8, 0, 4, 0);
                    inner.HorizontalAlignment = IsVertical ? HorizontalAlignment.Center : HorizontalAlignment.Stretch;
                }
                
                if (!IsVertical && t.TitleText != null) 
                    t.TitleText.MaxWidth = Math.Max(10, _tabW - 56);
            }

            UpdateScrollButtonState();
        }

        // スクロールボタンの有効・無効（薄くする）を更新
        private void UpdateScrollButtonState()
        {
            if (!_isScrollMode) return;
            double totalW = _tabs.Count * (_tabW + TabGap);
            double visible = TabBar.ActualWidth;

            // 左端より左にはスクロールできない
            double leftOpacity = _tabScrollOffset > 0 ? 1.0 : 0.3;
            // 右端より右にはスクロールできない
            double rightOpacity = (_tabScrollOffset + visible) < totalW ? 1.0 : 0.3;

            ((TextBlock)TabScrollLeft.Child).Opacity = leftOpacity;
            ((TextBlock)TabScrollRight.Child).Opacity = rightOpacity;
        }

        // スクロールを適用してタブを再描画
        private void ApplyScroll(double delta, bool animate = true)
        {
            double totalW = _tabs.Count * (_tabW + TabGap);
            double visible = TabBar.ActualWidth;
            double newOffset = Math.Max(0, Math.Min(totalW - visible, _tabScrollOffset + delta));
            if (Math.Abs(newOffset - _tabScrollOffset) < 0.1) return;
            _tabScrollOffset = newOffset;
            LayoutTabs(animate);
            UpdateScrollButtonState();
        }

        private void TabScrollLeft_Click(object sender, MouseButtonEventArgs e)
        {
            ApplyScroll(-(_tabW + TabGap) * 3); // 3タブ分左へ
        }

        private void TabScrollRight_Click(object sender, MouseButtonEventArgs e)
        {
            ApplyScroll((_tabW + TabGap) * 3);  // 3タブ分右へ
        }

        private void TabBar_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (!_isScrollMode) return;
            ApplyScroll(e.Delta > 0 ? -(_tabW + TabGap) : (_tabW + TabGap));
            e.Handled = true;
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
                double currentTabH = IsVertical ? 60.0 : TabH;
                double posVal = i * ((IsVertical ? currentTabH : _tabW) + TabGap) + 2 - _tabScrollOffset;
                
                if (IsVertical)
                {
                    Canvas.SetLeft(_tabs[i].Header, 0);
                    if (animate)
                        _tabs[i].Header.BeginAnimation(Canvas.TopProperty, new DoubleAnimation { To = posVal, Duration = dur, EasingFunction = ease });
                    else
                        Canvas.SetTop(_tabs[i].Header, posVal);
                }
                else
                {
                    Canvas.SetTop(_tabs[i].Header, 2);
                    if (animate)
                        _tabs[i].Header.BeginAnimation(Canvas.LeftProperty, new DoubleAnimation { To = posVal, Duration = dur, EasingFunction = ease });
                    else
                        Canvas.SetLeft(_tabs[i].Header, posVal);
                }
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
                if (Math.Abs(pos.X - _dragStartPoint.X) < DragThreshold && Math.Abs(pos.Y - _dragStartPoint.Y) < DragThreshold) return;
                _isDraggingTab = true;
                SelectTab(_tabs.IndexOf(_dragTab));
                Canvas.SetZIndex(_dragTab.Header, 999);
            }

            // ⭐ [要件4] エリア外判定（上下左右に 40px 以上ずれたら切り離し状態に移行）
            bool isOut = IsVertical ? (pos.X < -40 || pos.X > TabBar.ActualWidth + 40) : (pos.Y < -40 || pos.Y > TabBar.ActualHeight + 40);
            if (isOut)
            {
                if (_dragGhostWindow == null)
                {
                    _dragTab.Header.Visibility = Visibility.Hidden;
                    _dragGhostWindow = new Window
                    {
                        WindowStyle = WindowStyle.None, AllowsTransparency = true, Background = Brushes.Transparent, Topmost = true, ShowInTaskbar = false,
                        Width = _tabW, Height = TabH, Opacity = 0.8,
                        Content = new Border { Background = new VisualBrush(_dragTab.Header) { Stretch = Stretch.None } }
                    };
                    _dragGhostWindow.Show();
                }

                var screenPos = TabBar.PointToScreen(pos);
                _dragGhostWindow.Left = screenPos.X - (IsVertical ? 0 : _dragOffsetX);
                _dragGhostWindow.Top = screenPos.Y - (IsVertical ? _dragOffsetX : 10);
                return; // 通常のタブ挿入アニメーションは行わない
            }
            // エリア内に戻ってきたらゴーストを消して元に戻す
            if (_dragGhostWindow != null)
            {
                _dragGhostWindow.Close();
                _dragGhostWindow = null;
                _dragTab.Header.Visibility = Visibility.Visible;
            }

            // アニメーションのロックを解除して手動移動できるようにする
            if (IsVertical)
            {
                _dragTab.Header.BeginAnimation(Canvas.TopProperty, null);
            }
            else
            {
                _dragTab.Header.BeginAnimation(Canvas.LeftProperty, null);
            }

            double itemSize = IsVertical ? 60.0 : _tabW;
            double itemOffset = IsVertical ? pos.Y : pos.X;
            double minPos = 2;
            double maxPos = 2 + (_tabs.Count - 1) * (itemSize + TabGap);
            double newPos = Math.Max(minPos, Math.Min(maxPos, itemOffset - _dragOffsetX));

            if (IsVertical)
                Canvas.SetTop(_dragTab.Header, newPos);
            else
                Canvas.SetLeft(_dragTab.Header, newPos);

            int newIns = (int)((newPos - 2 + itemSize / 2) / (itemSize + TabGap));
            newIns = Math.Max(0, Math.Min(_tabs.Count - 1, newIns));
            if (newIns != _dragInsertIndex) { _dragInsertIndex = newIns; AnimateSlots(); }
        }

        private void TabBar_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_dragTab == null) return;
            _dragTab.Header.ReleaseMouseCapture();

            // ⭐ [要件5] ゴースト表示中（エリア外）でドロップされたら切り離す
            if (_dragGhostWindow != null)
            {
                _dragGhostWindow.Close();
                _dragGhostWindow = null;
                _dragTab.Header.Visibility = Visibility.Visible;
                
                var targetTab = _dragTab; // 参照を保持
                _dragTab = null;
                _isDraggingTab = false;
                
                DetachTab(targetTab); 
                e.Handled = true;
                return;
            }

            if (_isDraggingTab)
            {
                int from = _tabs.IndexOf(_dragTab), to = _dragInsertIndex;
                if (from != to) { _tabs.RemoveAt(from); _tabs.Insert(to, _dragTab); _selectedIndex = to; }
                foreach (var t in _tabs)
                {
                    t.Shift.BeginAnimation(TranslateTransform.XProperty, null);
                    t.Shift.BeginAnimation(TranslateTransform.YProperty, null);
                    t.Shift.X = 0;
                    t.Shift.Y = 0;
                    Canvas.SetZIndex(t.Header, 0);
                }
                LayoutTabs(animate: true);
            }
            else
            {
                SelectTab(_tabs.IndexOf(_dragTab));
            }

            _dragTab = null;
            _isDraggingTab = false;
            e.Handled = true;
        }

        private void AnimateSlots()
        {
            int dragFrom = _tabs.IndexOf(_dragTab);
            double itemSize = IsVertical ? 60.0 : _tabW;
            for (int i = 0; i < _tabs.Count; i++)
            {
                if (_tabs[i] == _dragTab) continue;
                double shift = 0;
                if (dragFrom < _dragInsertIndex && i > dragFrom && i <= _dragInsertIndex)
                    shift = -(itemSize + TabGap);
                else if (dragFrom > _dragInsertIndex && i >= _dragInsertIndex && i < dragFrom)
                    shift = itemSize + TabGap;
                
                var prop = IsVertical ? TranslateTransform.YProperty : TranslateTransform.XProperty;
                _tabs[i].Shift.BeginAnimation(prop,
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
        //  キャプションボタン & 追加機能ボタン
        // =====================================================================
        private void CaptionBtn_Down(object sender, MouseButtonEventArgs e) => e.Handled = true;
        private void BtnMinimize_Click(object sender, MouseButtonEventArgs e) { WindowState = WindowState.Minimized; e.Handled = true; }
        private void BtnMaximize_Click(object sender, MouseButtonEventArgs e) { 
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized; 
            e.Handled = true;
        }
        private void BtnClose_Click(object sender, MouseButtonEventArgs e) { Close(); e.Handled = true; }
        private void UpdateMaximizeIcon() { MaximizeIcon.Text = WindowState == WindowState.Maximized ? "❐" : "□"; }

        // [要件7, 8] すべてのウィンドウを集約する
        private void BtnAttachAll_Click(object sender, MouseButtonEventArgs e)
        {
            var myPid = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            NativeMethods.EnumWindows((hWnd, lParam) =>
            {
                if (!NativeMethods.IsWindowVisible(hWnd)) return true; // 非表示ウィンドウは除く
                
                NativeMethods.GetWindowThreadProcessId(hWnd, out uint pid);
                if (pid == myPid) return true; // 自分自身は除く

                // タイトルが無いウィンドウは実体として表示されないことが多い
                var title = GetTitle(hWnd);
                if (string.IsNullOrWhiteSpace(title)) return true;

                // 通常のウィンドウアプリはタイトルバー（WS_CAPTION）を持っている
                var style = NativeMethods.GetWindowLongPtr(hWnd, NativeMethods.GWL_STYLE).ToInt64();
                if ((style & NativeMethods.WS_CAPTION) == 0) return true;

                // アタッチ処理を通す (UWP除外機能は AttachWindow 内に実装済み)
                AttachWindow(hWnd);
                return true;
            }, IntPtr.Zero);
            e.Handled = true;
        }

        // [要件9, 10, 11] 設定メニューとレイアウトリアルタイム切り替え
        private void ChangeTabPlacement(string placement)
        {
            if (_tabPlacement == placement) return;
            _tabPlacement = placement;

            // リセット
            RowTop.Height = new GridLength(30); // 上以外ならキャプションボタン列を細く
            RowBottom.Height = new GridLength(0);
            ColLeft.Width = new GridLength(0);
            ColRight.Width = new GridLength(0);
            TabBarContainer.Margin = new Thickness(0);
            
            // タブの方向切り替え
            TabScrollLeft.SetValue(Grid.RowProperty, 0);
            TabScrollLeft.SetValue(Grid.ColumnProperty, 0);
            TabScrollRight.SetValue(Grid.RowProperty, 0);
            TabScrollRight.SetValue(Grid.ColumnProperty, 0);

            switch (placement)
            {
                case "上側 (Top)":
                    RowTop.Height = new GridLength(42); // 上配置時はキャプション列と同じ行にするため広く
                    Grid.SetRow(TabBarContainer, 0);
                    Grid.SetColumn(TabBarContainer, 0);
                    Grid.SetColumnSpan(TabBarContainer, 3);
                    Grid.SetRowSpan(TabBarContainer, 1);
                    TabBarContainer.Margin = new Thickness(0, 0, 180, 0); // キャプションボタンを避ける
                    
                    TabScrollLeft.Width = 28; TabScrollLeft.Height = double.NaN;
                    TabScrollLeft.HorizontalAlignment = HorizontalAlignment.Left;
                    TabScrollLeft.VerticalAlignment = VerticalAlignment.Stretch;
                    TabScrollLeft.Margin = new Thickness(0);
                    
                    TabScrollRight.Width = 28; TabScrollRight.Height = double.NaN;
                    TabScrollRight.HorizontalAlignment = HorizontalAlignment.Left;
                    TabScrollRight.VerticalAlignment = VerticalAlignment.Stretch;
                    TabScrollRight.Margin = new Thickness(28, 0, 0, 0);
                    break;
                case "下側 (Bottom)":
                    RowBottom.Height = new GridLength(42);
                    Grid.SetRow(TabBarContainer, 2);
                    Grid.SetColumn(TabBarContainer, 0);
                    Grid.SetColumnSpan(TabBarContainer, 3);
                    Grid.SetRowSpan(TabBarContainer, 1);
                    
                    TabScrollLeft.Width = 28; TabScrollLeft.Height = double.NaN;
                    TabScrollLeft.HorizontalAlignment = HorizontalAlignment.Left;
                    TabScrollLeft.VerticalAlignment = VerticalAlignment.Stretch;
                    TabScrollLeft.Margin = new Thickness(0);
                    
                    TabScrollRight.Width = 28; TabScrollRight.Height = double.NaN;
                    TabScrollRight.HorizontalAlignment = HorizontalAlignment.Left;
                    TabScrollRight.VerticalAlignment = VerticalAlignment.Stretch;
                    TabScrollRight.Margin = new Thickness(28, 0, 0, 0);
                    break;
                case "左側 (Left)":
                    ColLeft.Width = new GridLength(60);
                    Grid.SetRow(TabBarContainer, 1);
                    Grid.SetColumn(TabBarContainer, 0);
                    Grid.SetColumnSpan(TabBarContainer, 1);
                    Grid.SetRowSpan(TabBarContainer, 1);
                    
                    TabScrollLeft.Width = double.NaN; TabScrollLeft.Height = 28;
                    TabScrollLeft.HorizontalAlignment = HorizontalAlignment.Stretch;
                    TabScrollLeft.VerticalAlignment = VerticalAlignment.Top;
                    TabScrollLeft.Margin = new Thickness(0);
                    
                    TabScrollRight.Width = double.NaN; TabScrollRight.Height = 28;
                    TabScrollRight.HorizontalAlignment = HorizontalAlignment.Stretch;
                    TabScrollRight.VerticalAlignment = VerticalAlignment.Top;
                    TabScrollRight.Margin = new Thickness(0, 28, 0, 0);
                    break;
                case "右側 (Right)":
                    ColRight.Width = new GridLength(60);
                    Grid.SetRow(TabBarContainer, 1);
                    Grid.SetColumn(TabBarContainer, 2);
                    Grid.SetColumnSpan(TabBarContainer, 1);
                    Grid.SetRowSpan(TabBarContainer, 1);
                    
                    TabScrollLeft.Width = double.NaN; TabScrollLeft.Height = 28;
                    TabScrollLeft.HorizontalAlignment = HorizontalAlignment.Stretch;
                    TabScrollLeft.VerticalAlignment = VerticalAlignment.Top;
                    TabScrollLeft.Margin = new Thickness(0);
                    
                    TabScrollRight.Width = double.NaN; TabScrollRight.Height = 28;
                    TabScrollRight.HorizontalAlignment = HorizontalAlignment.Stretch;
                    TabScrollRight.VerticalAlignment = VerticalAlignment.Top;
                    TabScrollRight.Margin = new Thickness(0, 28, 0, 0);
                    break;
            }

            RecalcTabWidths(false);
            LayoutTabs(false);
            SyncActiveWindow();
        }

        private void BtnSettings_Click(object sender, MouseButtonEventArgs e)
        {
            var ctx = new ContextMenu();
            var placements = new[] { "上側 (Top)", "下側 (Bottom)", "左側 (Left)", "右側 (Right)" };
            foreach (var pos in placements)
            {
                var mi = new MenuItem { Header = pos, IsChecked = pos == _tabPlacement };
                mi.Click += (s, ev) => ChangeTabPlacement(pos);
                ctx.Items.Add(mi);
            }
            ctx.PlacementTarget = BtnSettings;
            ctx.IsOpen = true;
            e.Handled = true;
        }

        private void CaptionBtn_Enter(object sender, MouseEventArgs e) => ((Border)sender).Background = new SolidColorBrush(Color.FromArgb(0x30, 0xff, 0xff, 0xff));
        private void CloseBtn_Enter(object sender, MouseEventArgs e) => ((Border)sender).Background = new SolidColorBrush(Color.FromArgb(0xe0, 0xe8, 0x11, 0x23));
        private void CaptionBtn_Leave(object sender, MouseEventArgs e) => ((Border)sender).Background = Brushes.Transparent;

        // =====================================================================
        //  操作 
        // =====================================================================
        private async void CloseApp(TabData tab)
        {
            var hWnd = tab.HWnd;
            _tabs.Remove(tab);
            TabBar.Children.Remove(tab.Header);
            RestoreStyles(hWnd);
            NativeMethods.PostMessage(hWnd, NativeMethods.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            
            // ⭐ [要件6] 描画不具合対策：閉じる瞬間に一瞬最小化して復元させ、再描画を強制する
            NativeMethods.ShowWindow(hWnd, NativeMethods.SW_MINIMIZE);
            await System.Threading.Tasks.Task.Delay(10);
            NativeMethods.ShowWindow(hWnd, NativeMethods.SW_RESTORE);

            if (_tabs.Count > 0) SelectTab(0); else DropHintText.Visibility = Visibility.Visible;
            RecalcTabWidths();
            LayoutTabs();
        }

        private async void DetachTab(TabData tab)
        {
            var hWnd = tab.HWnd;
            _tabs.Remove(tab);
            TabBar.Children.Remove(tab.Header);
            RestoreStyles(hWnd);
            
            // ⭐ [要件6] 描画不具合対策：切り離し時も再描画を強制する
            NativeMethods.ShowWindow(hWnd, NativeMethods.SW_MINIMIZE);
            await System.Threading.Tasks.Task.Delay(10);
            NativeMethods.ShowWindow(hWnd, NativeMethods.SW_RESTORE);

            NativeMethods.GetCursorPos(out var pt);
            NativeMethods.SetWindowPos(hWnd, IntPtr.Zero, pt.X - 100, pt.Y - 10, 800, 600, NativeMethods.SWP_SHOWWINDOW | NativeMethods.SWP_FRAMECHANGED);
            if (_tabs.Count > 0)
            {
                SelectTab(0);
            }
            else
            {
                DropHintText.Visibility = Visibility.Visible;
                DropHintTextCenter.Visibility = Visibility.Visible;
            }
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
            foreach (var t in _tabs) 
            { 
                RestoreStyles(t.HWnd); 
                // アプリ終了時も同様に対応
                NativeMethods.ShowWindow(t.HWnd, NativeMethods.SW_MINIMIZE);
                System.Threading.Thread.Sleep(10);
                NativeMethods.ShowWindow(t.HWnd, NativeMethods.SW_RESTORE);
            }
        }

        private async System.Threading.Tasks.Task<ImageSource> GetAppIconAsync(IntPtr hWnd)
        {
            return await System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    NativeMethods.GetWindowThreadProcessId(hWnd, out uint pid);
                    IntPtr hProc = NativeMethods.OpenProcess(NativeMethods.PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
                    if (hProc != IntPtr.Zero)
                    {
                        var sb = new StringBuilder(1024);
                        int size = sb.Capacity;
                        if (NativeMethods.QueryFullProcessImageName(hProc, 0, sb, ref size))
                        {
                            NativeMethods.CloseHandle(hProc);
                            var path = sb.ToString();
                            var icon = System.Drawing.Icon.ExtractAssociatedIcon(path);
                            if (icon != null)
                            {
                                // WPF で安全に扱えるように、GDIとの縁を切って完全にメモリ上にコピーする
                                using (var bmp = icon.ToBitmap())
                                {
                                    using (var ms = new System.IO.MemoryStream())
                                    {
                                        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                        var bytes = ms.ToArray();
                                        return Dispatcher.Invoke(() =>
                                        {
                                            var bi = new System.Windows.Media.Imaging.BitmapImage();
                                            bi.BeginInit();
                                            bi.StreamSource = new System.IO.MemoryStream(bytes);
                                            bi.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                                            bi.EndInit();
                                            bi.Freeze();
                                            return bi;
                                        });
                                    }
                                }
                            }
                        }
                        NativeMethods.CloseHandle(hProc);
                    }
                }
                catch { }
                return null;
            });
        }

        private static string GetTitle(IntPtr hWnd)
        {
            int len = NativeMethods.GetWindowTextLength(hWnd);
            var sb = new StringBuilder(len + 1);
            NativeMethods.GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private static string GetClassName(IntPtr hWnd)
        {
            var sb = new StringBuilder(256);
            NativeMethods.GetClassName(hWnd, sb, sb.Capacity);
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
            var oldCenter = DropHintTextCenter.Text; DropHintTextCenter.Text = msg; DropHintTextCenter.Visibility = Visibility.Visible;
            
            await System.Threading.Tasks.Task.Delay(3000);
            
            DropHintText.Text = old; 
            DropHintText.Visibility = _tabs.Count > 0 ? Visibility.Collapsed : Visibility.Visible;
            
            DropHintTextCenter.Text = oldCenter;
            DropHintTextCenter.Visibility = _tabs.Count > 0 ? Visibility.Collapsed : Visibility.Visible;
        }
    }
}
