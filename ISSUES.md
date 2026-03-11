# Tabify - Task / Issue 管理

## 凡例

- 🔴 **Critical**: クラッシュや機能不全に直結する問題
- 🟡 **Warning**: 動作はするが意図しない挙動・将来的なバグの温床
- 🟢 **Enhancement**: パフォーマンスや保守性の改善提案
- ⬜ **未着手** / 🔲 **対応中** / ✅ **完了**

---

## 🔴 Critical

### ✅ #C-001: BuildTabHeader の子要素型不一致によるスタイル適用漏れ (v1.0.26 で修正)

- **場所**: `MainWindow.xaml.cs` - `BuildTabHeader`, `ApplyTheme`, `RecalcTabWidths`
- **内容**: `BuildTabHeader` では `InnerBorder.Child` に `StackPanel` を設定しているが、`ApplyTheme` や `RecalcTabWidths` では `inner.Child is Grid stack` と `Grid` にキャストしている。`is` パターンマッチにより例外は発生しないが、テーマ切替時の閉じるボタン色反映、左右配置時のアイコンサイズ変更などがサイレントにスキップされている。
- **影響**: ダーク↔ライト切替時にタブ内の `✕` ボタン色が更新されない。左右配置時のアイコン/閉じるボタン表示制御が効かない。
- **対策案**: `StackPanel` → `Grid` に統一するか、キャスト先を `StackPanel` に修正する。理想的にはタブUI構造を UserControl または DataTemplate に切り出して型安全にする。

---

## 🟡 Warning

### ✅ #W-001: _syncTimer が常時稼働（CPU 空転） (v1.0.27 で修正)

- **場所**: `MainWindow.xaml.cs` - コンストラクタ / `_syncTimer`
- **内容**: 15ms 間隔の `DispatcherTimer` がタブ 0 個の待機状態でも回り続ける。`SyncActiveWindow` 内で早期リターンはしているが、タイマーコールバック自体の発火コストが常時発生する。
- **影響**: アイドル時の CPU 使用率がわずかに上昇。ノート PC のバッテリー消費に影響する可能性。
- **対策案**: タブが 0 個になったら `_syncTimer.Stop()`、タブ追加時に `_syncTimer.Start()` とする。

### ✅ #W-002: ForceAttach で Win32 スタイル変更が行われない (v1.0.29 で修正、タブ移動の問題は別途調査)

- **場所**: `MainWindow.xaml.cs` - `ForceAttach`
- **内容**: 別の Tabify ウィンドウからタブを受け取る `ForceAttach` では、`_origStyles` / `_origParents` への登録や `WS_CAPTION` / `WS_THICKFRAME` の除去が行われていない。元の Tabify 側で既にスタイル変更済みの前提だが、`GWLP_HWNDPARENT` の付け替え（新しいコンテナへの紐付け）も未実施。
- **影響**: ウィンドウ間でタブを移動した後、Z オーダーや表示/非表示の制御が不安定になる可能性がある。
- **対策案**: `ForceAttach` 内でも `GWLP_HWNDPARENT` を新しいコンテナの HWND に再設定する。

### ✅ #W-003: タブタイトルが動的に更新されない (v1.0.32 で修正)

- **場所**: `MainWindow.xaml.cs` - `TabData.Title`, `_syncTimer`
- **内容**: タブのタイトルはアタッチ時に一度だけ `GetTitle(hWnd)` で取得されるが、対象アプリがタイトルを変更しても（例: ブラウザのページ遷移、エディタのファイル切替）タブ表示に反映されない。
- **影響**: ユーザーがどのタブがどのコンテンツかを判別できなくなる。
- **対策案**: `_syncTimer` のコールバック内で定期的に `GetWindowText` を呼び、変化があればタブの `TitleText.Text` と `ToolTip` を更新する。

### ✅ #W-004: 対象ウィンドウの終了（外部からの閉じる操作）を検知していない (v1.0.33 で修正)

- **場所**: `MainWindow.xaml.cs` 全体
- **内容**: ユーザーがタスクマネージャーや Alt+F4 等で対象アプリを直接終了した場合、Tabify 側のタブは残り続けるが、対象の HWND は無効になる。`SyncActiveWindow` で無効なハンドルに `SetWindowPos` を呼び続ける。
- **影響**: 無効ハンドルへの操作による予期しない挙動。空のタブが残留する。
- **対策案**: `_syncTimer` 内で `IsWindow(hWnd)` をチェックし、無効ならタブを自動除去する。`NativeMethods` に `IsWindow` の P/Invoke を追加する。

### ⬜ #W-005: WinEventHook のイベント範囲が狭い

- **場所**: `MainWindow.xaml.cs` - コンストラクタ
- **内容**: `SetWinEventHook` が `EVENT_SYSTEM_MOVESIZESTART` ～ `EVENT_SYSTEM_MOVESIZEEND` の範囲のみをフックしている。対象ウィンドウの破棄（`EVENT_OBJECT_DESTROY`）やフォーカス変更（`EVENT_SYSTEM_FOREGROUND`）などは検知できない。
- **影響**: #W-004 と関連。外部終了の即時検知ができない。
- **対策案**: `EVENT_OBJECT_DESTROY` (0x8001) のフックを追加し、管理中の HWND が破棄されたらタブを自動除去する。

### ✅ #W-006: ChangeTabPlacement で RowSpan/ColumnSpan がリセットされない (v1.0.34 で修正)

- **場所**: `MainWindow.xaml.cs` - `ChangeTabPlacement`
- **内容**: 例えば「上側」→「左側」→「上側」と切り替えた場合、`Grid.SetRowSpan` が左側設定の `1` のまま残り、上側ケースで `RowSpan` の再設定呼び出しがない。
- **影響**: 特定の切り替え順序でタブバーのレイアウトが崩れる可能性。
- **対策案**: `ChangeTabPlacement` の冒頭で `RowSpan=1, ColumnSpan=1` にリセットしてから各配置の値を設定する。

### ⬜ #W-007: DetachTab 後のウィンドウサイズがハードコード

- **場所**: `MainWindow.xaml.cs` - `DetachTab`
- **内容**: デタッチ後のウィンドウサイズが `800x600` にハードコードされている。元のウィンドウサイズ情報（アタッチ前の `RECT`）が保存されていないため、元の大きさに戻せない。
- **影響**: 巨大なモニタで小さなウィンドウとして復元される、または小さなモニタではみ出す可能性。
- **対策案**: アタッチ時に `GetWindowRect` で元のサイズを `TabData` に保存し、デタッチ時に復元する。

---

### ⬜ #W-008: Alt+F4 が Tabify 自体に届いてしまう

- **場所**: `MainWindow.xaml.cs` 全体
- **内容**: タブ内のアプリにフォーカスがあるように見えても、実際のフォアグラウンドウィンドウは Tabify であるため、Alt+F4 を押すと対象アプリではなく Tabify 自体が閉じてしまう。
- **影響**: ユーザーがタブ内のアプリを Alt+F4 で閉じようとすると、意図せず Tabify ごと終了する。
- **対策案**: `WndProc` で `WM_CLOSE` や `WM_SYSKEYDOWN`（Alt+F4）をフックし、タブが存在する場合はアクティブなタブの対象ウィンドウに `WM_CLOSE` を転送する。または Tabify 終了前に確認ダイアログを表示する。

---

## 🟢 Enhancement

### ⬜ #E-001: タブ UI 構造の UserControl / DataTemplate 化

- **場所**: `MainWindow.xaml.cs` - `BuildTabHeader`
- **内容**: タブの UI が C# コードビハインドで手動構築（`new Border`, `new StackPanel` 等）されており、変更時の影響範囲が広く可読性が低い。XAML の `DataTemplate` または `UserControl` に切り出すことで保守性が大幅に向上する。
- **優先度**: 中（リファクタリング）

### ⬜ #E-002: タブドラッグ中のスクロール自動追従

- **場所**: `MainWindow.xaml.cs` - `TabBar_MouseMove`
- **内容**: スクロールモード中にタブをドラッグしてスクロール端に到達しても、自動的にスクロールされない。大量のタブがある場合、端のタブを反対側へ移動するのが困難。
- **優先度**: 低

### ⬜ #E-003: マルチモニター DPI 混在環境への対応

- **場所**: `MainWindow.xaml.cs` - `SyncActiveWindow`
- **内容**: `ContentArea.PointToScreen` と `TransformToDevice` で座標変換しているが、異なる DPI のモニター間でウィンドウを移動した際に正しく追従するか未検証。Per-Monitor DPI Awareness V2 の宣言（`app.manifest`）も未設定。
- **優先度**: 中

### ⬜ #E-004: エラーログ出力先のハードコード

- **場所**: `MainWindow.xaml.cs` - `AttachWindow` の catch 節
- **内容**: `C:\Users\harak\Documents\MyGit\work\20260301-likeGroupy\attach_error.log` という開発者固有のパスにハードコードされている。他の環境では書き込み失敗する。
- **優先度**: 高（リリース前に修正必須）
- **対策案**: `Environment.GetFolderPath` や `AppDomain.CurrentDomain.BaseDirectory` を使用する。またはリリースビルドではファイル書き込みを無効化する。

### ⬜ #E-005: README.md のバージョンバッジが手動更新

- **場所**: `README.md`
- **内容**: `version-1.0.17` のバッジが静的に記述されており、`publish.ps1` で `.csproj` のバージョンを上げても README には反映されない（現在 csproj は `1.0.24`）。
- **優先度**: 低
- **対策案**: `publish.ps1` 内で README のバッジも自動更新するか、GitHub Actions で動的バッジを生成する。

---

## 変更履歴

| 日付 | 内容 |
|------|------|
| 2026-03-10 | 初版作成（12件: C×1, W×7, E×5） |
