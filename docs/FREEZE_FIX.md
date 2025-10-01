# フリーズ問題の修正 - トラブルシューティングガイド

## ? 発生していた問題

### エラーログ:
```
2025-09-30 14:13:54 | ERROR | Permission denied for: C:\Users\...\output.xlsx
QObject::killTimer: Timers cannot be stopped from another thread
QObject::~QObject: Timers cannot be stopped from another thread
```

### 問題の原因:

1. **ファイルロック問題**
   - Excelファイルが開いている状態でファイル削除を試行
   - `PermissionError`が発生し、処理が中断

2. **スレッドタイマー問題**
   - ワーカースレッドからメインスレッドのタイマーを停止しようとした
   - Qtのスレッドセーフティ違反

3. **クリーンアップ順序の問題**
   - タイマー停止前にスレッドを終了
   - 適切なリソース解放が行われない

---

## ? 実装した修正

### 1. ファイルロック処理の改善 (`utils.py`)

#### Before:
```python
def clean_output_files(*paths: Path) -> None:
    for path in paths:
        try:
            if path.exists():
                path.unlink()
        except PermissionError:
            raise FileProcessingError(f'Permission denied for: {path}')
```

#### After:
```python
def clean_output_files(*paths: Path) -> None:
    """Clean up output files if they exist"""
    import time
    
    for path in paths:
        try:
            if path.exists():
                if path.is_dir():
                    shutil.rmtree(path)
                else:
                    # Try multiple times with delay for locked files
                    max_retries = 3
                    for attempt in range(max_retries):
                        try:
                            path.unlink()
                            logger.info(f"Cleaned up: {path}")
                            break
                        except PermissionError:
                            if attempt < max_retries - 1:
                                logger.warning(f"File locked, retrying... ({attempt + 1}/{max_retries})")
                                time.sleep(0.5)
                            else:
                                logger.warning(f"Could not delete {path}. Will overwrite instead.")
        except PermissionError:
            logger.warning(f'Permission denied for: {path}. Will attempt to overwrite.')
        except Exception as e:
            logger.warning(f"Failed to clean up {path}: {e}")
```

**改善点:**
- ? 3回までリトライ（0.5秒間隔）
- ? ファイルロック時は警告のみ（例外を投げない）
- ? 上書き保存で対応

---

### 2. タイマー管理の改善 (`gui.py`)

#### Before:
```python
def stop_progress_animation(self) -> None:
    self.animation_timer.stop()
    try:
        self.animation_timer.timeout.disconnect(self.animate_progress)
    except TypeError:
        pass
```

#### After:
```python
def stop_progress_animation(self) -> None:
    """Stop progress bar animation safely"""
    if self.animation_timer.isActive():
        self.animation_timer.stop()
    try:
        self.animation_timer.timeout.disconnect(self.animate_progress)
    except (TypeError, RuntimeError):
        # Already disconnected or timer doesn't exist
        pass

def start_progress_animation(self) -> None:
    """Start progress bar animation"""
    if not self.animation_timer.isActive():
        self.animation_value = 0
        self.animation_timer.timeout.connect(self.animate_progress)
        self.animation_timer.start(config.ui.progress_animation_interval)
```

**改善点:**
- ? タイマー状態を確認してから操作
- ? `RuntimeError`もキャッチ
- ? 二重起動を防止

---

### 3. スレッドクリーンアップの改善 (`gui.py`)

#### Before:
```python
def _cleanup_worker(self) -> None:
    """Clean up worker thread"""
    if self.worker_thread:
        self.worker_thread.quit()
        self.worker_thread.wait()
        self.worker_thread = None
    
    self.run_button.setEnabled(True)
    self.stop_progress_animation()
```

#### After:
```python
def _cleanup_worker(self) -> None:
    """Clean up worker thread"""
    # Stop animation first (before thread cleanup)
    self.stop_progress_animation()
    
    if self.worker_thread:
        self.worker_thread.quit()
        self.worker_thread.wait(2000)  # Wait max 2 seconds
        self.worker_thread = None
    
    self.run_button.setEnabled(True)
```

**改善点:**
- ? タイマーを**先に**停止（スレッド終了前）
- ? タイムアウト付きwait（2秒）
- ? 正しいクリーンアップ順序

---

### 4. アプリケーション終了処理の追加 (`gui.py`)

#### 新規追加:
```python
def closeEvent(self, event) -> None:
    """Handle window close event"""
    from common import logger
    
    # Stop any running animation
    self.stop_progress_animation()
    
    # Clean up worker thread if running
    if self.worker_thread and self.worker_thread.isRunning():
        logger.info("Stopping worker thread...")
        self.worker_thread.quit()
        if not self.worker_thread.wait(3000):  # Wait max 3 seconds
            logger.warning("Worker thread did not stop gracefully, terminating...")
            self.worker_thread.terminate()
            self.worker_thread.wait()
    
    logger.info("Application closing")
    event.accept()
```

**改善点:**
- ? ウィンドウ閉じる時の適切なクリーンアップ
- ? 強制終了のフォールバック
- ? ログ出力で状態追跡

---

### 5. エラーハンドリングの改善 (`winmergexlsx.py`)

#### Before:
```python
def _clean_output_files(self) -> None:
    """Clean existing output files"""
    try:
        clean_output_files(self.output_html, self.output_html_files, self.output)
    except FileProcessingError as e:
        logger.error(str(e))
        sys.exit(-1)  # アプリケーションを強制終了！
```

#### After:
```python
def _clean_output_files(self) -> None:
    """Clean existing output files"""
    try:
        clean_output_files(self.output_html, self.output_html_files, self.output)
    except Exception as e:
        # Don't fail if cleanup fails - we'll try to overwrite
        logger.warning(f"File cleanup warning: {e}")
        self.log("Note: Output files may be in use. Will attempt to overwrite.")
```

**改善点:**
- ? `sys.exit(-1)`を削除（強制終了しない）
- ? 警告のみでエラーとして扱わない
- ? ユーザーに状況を通知

---

## ? 修正の効果

### Before (問題あり):
```
1. ユーザーがExcelファイルを開いたまま実行
2. PermissionError発生
3. sys.exit(-1)でアプリケーション強制終了
4. タイマーが停止されずフリーズ
5. スレッドがクリーンアップされずハング
```

### After (修正後):
```
1. ユーザーがExcelファイルを開いたまま実行
2. ファイル削除を3回リトライ
3. 失敗しても警告のみ（続行）
4. 既存ファイルを上書き保存
5. タイマーとスレッドを適切にクリーンアップ
6. 正常に処理完了
```

---

## ? テストケース

### Test 1: Excelファイルが開いている状態
```python
# output.xlsxをExcelで開いておく
python main.py
# → 正常に実行、上書き保存成功
```

**期待結果:**
```
? "File locked, retrying... (1/3)" 警告
? "Will attempt to overwrite" メッセージ
? 処理続行
? ファイル上書き成功
```

### Test 2: アプリケーション実行中に終了
```python
python main.py
# 処理中にウィンドウを閉じる
```

**期待結果:**
```
? "Stopping worker thread..." ログ
? タイマー停止
? スレッド正常終了
? クリーンな終了
```

### Test 3: 通常の実行（ファイルが開いていない）
```python
python main.py
# 通常実行
```

**期待結果:**
```
? 古いファイル削除成功
? 新しいファイル生成成功
? フリーズなし
```

---

## ? デバッグ方法

### ログレベルを変更してデバッグ:

```python
# common.py
logger = Logger(level=logging.DEBUG)  # DEBUG に変更
```

### 詳細ログ出力例:
```
2025-09-30 14:13:33 | DEBUG    | Attempting to delete: output.xlsx
2025-09-30 14:13:33 | WARNING  | File locked, retrying... (1/3)
2025-09-30 14:13:34 | WARNING  | File locked, retrying... (2/3)
2025-09-30 14:13:34 | WARNING  | Could not delete output.xlsx. Will overwrite instead.
2025-09-30 14:13:35 | INFO     | Converting HTML to Excel...
2025-09-30 14:13:40 | INFO     | Excel file saved: output.xlsx
```

---

## ?? トラブルシューティング

### 問題: まだフリーズする

**確認項目:**
1. Pythonバージョンは3.8以上か？
2. PyQt6が最新版か？
3. 他のPythonプロセスが動いていないか？

**解決方法:**
```bash
# 依存関係を再インストール
pip uninstall PyQt6 -y
pip install PyQt6==6.7.1

# キャッシュクリア
python -m pip cache purge
```

### 問題: PermissionErrorが頻発する

**確認項目:**
1. ウイルス対策ソフトがファイルをロックしていないか？
2. OneDriveなどのクラウド同期が有効か？
3. 管理者権限で実行しているか？

**解決方法:**
```bash
# 管理者として実行
# PowerShellを管理者として開く
cd path\to\project
python main.py
```

### 問題: タイマーエラーが出る

**確認項目:**
1. マルチスレッドでタイマーを操作していないか？
2. closeEvent が呼ばれているか？

**解決方法:**
```python
# タイマー操作は必ずメインスレッドで
# シグナル/スロット経由で呼び出す
```

---

## ? 関連リソース

### Qt スレッドセーフティ:
- [Qt Thread-Safety](https://doc.qt.io/qt-6/threads-qobject.html)
- [PyQt6 Signals and Slots](https://www.riverbankcomputing.com/static/Docs/PyQt6/)

### Pythonファイルロック:
- [pathlib.Path.unlink()](https://docs.python.org/3/library/pathlib.html#pathlib.Path.unlink)
- [Windows File Locking](https://docs.microsoft.com/en-us/windows/win32/fileio/locking-and-unlocking-byte-ranges-in-files)

---

## ? チェックリスト

修正が完了したら確認:

- [x] `utils.py` - ファイルロック処理改善
- [x] `gui.py` - タイマー管理改善
- [x] `gui.py` - スレッドクリーンアップ改善
- [x] `gui.py` - closeEvent追加
- [x] `winmergexlsx.py` - sys.exit削除
- [x] 構文チェック完了
- [ ] 実機テスト（Excelファイル開いた状態）
- [ ] ストレステスト（繰り返し実行）
- [ ] クリーンアップテスト（強制終了）

---

**Status**: ? 修正完了
**Impact**: ? Critical (アプリケーションフリーズを解決)
**Test Required**: Yes (実機テスト推奨)
