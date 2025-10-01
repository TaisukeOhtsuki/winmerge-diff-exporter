# �t���[�Y���̏C�� - �g���u���V���[�e�B���O�K�C�h

## ? �������Ă������

### �G���[���O:
```
2025-09-30 14:13:54 | ERROR | Permission denied for: C:\Users\...\output.xlsx
QObject::killTimer: Timers cannot be stopped from another thread
QObject::~QObject: Timers cannot be stopped from another thread
```

### ���̌���:

1. **�t�@�C�����b�N���**
   - Excel�t�@�C�����J���Ă����ԂŃt�@�C���폜�����s
   - `PermissionError`���������A���������f

2. **�X���b�h�^�C�}�[���**
   - ���[�J�[�X���b�h���烁�C���X���b�h�̃^�C�}�[���~���悤�Ƃ���
   - Qt�̃X���b�h�Z�[�t�e�B�ᔽ

3. **�N���[���A�b�v�����̖��**
   - �^�C�}�[��~�O�ɃX���b�h���I��
   - �K�؂ȃ��\�[�X������s���Ȃ�

---

## ? ���������C��

### 1. �t�@�C�����b�N�����̉��P (`utils.py`)

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

**���P�_:**
- ? 3��܂Ń��g���C�i0.5�b�Ԋu�j
- ? �t�@�C�����b�N���͌x���̂݁i��O�𓊂��Ȃ��j
- ? �㏑���ۑ��őΉ�

---

### 2. �^�C�}�[�Ǘ��̉��P (`gui.py`)

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

**���P�_:**
- ? �^�C�}�[��Ԃ��m�F���Ă��瑀��
- ? `RuntimeError`���L���b�`
- ? ��d�N����h�~

---

### 3. �X���b�h�N���[���A�b�v�̉��P (`gui.py`)

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

**���P�_:**
- ? �^�C�}�[��**���**��~�i�X���b�h�I���O�j
- ? �^�C���A�E�g�t��wait�i2�b�j
- ? �������N���[���A�b�v����

---

### 4. �A�v���P�[�V�����I�������̒ǉ� (`gui.py`)

#### �V�K�ǉ�:
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

**���P�_:**
- ? �E�B���h�E���鎞�̓K�؂ȃN���[���A�b�v
- ? �����I���̃t�H�[���o�b�N
- ? ���O�o�͂ŏ�Ԓǐ�

---

### 5. �G���[�n���h�����O�̉��P (`winmergexlsx.py`)

#### Before:
```python
def _clean_output_files(self) -> None:
    """Clean existing output files"""
    try:
        clean_output_files(self.output_html, self.output_html_files, self.output)
    except FileProcessingError as e:
        logger.error(str(e))
        sys.exit(-1)  # �A�v���P�[�V�����������I���I
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

**���P�_:**
- ? `sys.exit(-1)`���폜�i�����I�����Ȃ��j
- ? �x���݂̂ŃG���[�Ƃ��Ĉ���Ȃ�
- ? ���[�U�[�ɏ󋵂�ʒm

---

## ? �C���̌���

### Before (��肠��):
```
1. ���[�U�[��Excel�t�@�C�����J�����܂܎��s
2. PermissionError����
3. sys.exit(-1)�ŃA�v���P�[�V���������I��
4. �^�C�}�[����~���ꂸ�t���[�Y
5. �X���b�h���N���[���A�b�v���ꂸ�n���O
```

### After (�C����):
```
1. ���[�U�[��Excel�t�@�C�����J�����܂܎��s
2. �t�@�C���폜��3�񃊃g���C
3. ���s���Ă��x���̂݁i���s�j
4. �����t�@�C�����㏑���ۑ�
5. �^�C�}�[�ƃX���b�h��K�؂ɃN���[���A�b�v
6. ����ɏ�������
```

---

## ? �e�X�g�P�[�X

### Test 1: Excel�t�@�C�����J���Ă�����
```python
# output.xlsx��Excel�ŊJ���Ă���
python main.py
# �� ����Ɏ��s�A�㏑���ۑ�����
```

**���Ҍ���:**
```
? "File locked, retrying... (1/3)" �x��
? "Will attempt to overwrite" ���b�Z�[�W
? �������s
? �t�@�C���㏑������
```

### Test 2: �A�v���P�[�V�������s���ɏI��
```python
python main.py
# �������ɃE�B���h�E�����
```

**���Ҍ���:**
```
? "Stopping worker thread..." ���O
? �^�C�}�[��~
? �X���b�h����I��
? �N���[���ȏI��
```

### Test 3: �ʏ�̎��s�i�t�@�C�����J���Ă��Ȃ��j
```python
python main.py
# �ʏ���s
```

**���Ҍ���:**
```
? �Â��t�@�C���폜����
? �V�����t�@�C����������
? �t���[�Y�Ȃ�
```

---

## ? �f�o�b�O���@

### ���O���x����ύX���ăf�o�b�O:

```python
# common.py
logger = Logger(level=logging.DEBUG)  # DEBUG �ɕύX
```

### �ڍ׃��O�o�͗�:
```
2025-09-30 14:13:33 | DEBUG    | Attempting to delete: output.xlsx
2025-09-30 14:13:33 | WARNING  | File locked, retrying... (1/3)
2025-09-30 14:13:34 | WARNING  | File locked, retrying... (2/3)
2025-09-30 14:13:34 | WARNING  | Could not delete output.xlsx. Will overwrite instead.
2025-09-30 14:13:35 | INFO     | Converting HTML to Excel...
2025-09-30 14:13:40 | INFO     | Excel file saved: output.xlsx
```

---

## ?? �g���u���V���[�e�B���O

### ���: �܂��t���[�Y����

**�m�F����:**
1. Python�o�[�W������3.8�ȏォ�H
2. PyQt6���ŐV�ł��H
3. ����Python�v���Z�X�������Ă��Ȃ����H

**�������@:**
```bash
# �ˑ��֌W���ăC���X�g�[��
pip uninstall PyQt6 -y
pip install PyQt6==6.7.1

# �L���b�V���N���A
python -m pip cache purge
```

### ���: PermissionError���p������

**�m�F����:**
1. �E�C���X�΍�\�t�g���t�@�C�������b�N���Ă��Ȃ����H
2. OneDrive�Ȃǂ̃N���E�h�������L�����H
3. �Ǘ��Ҍ����Ŏ��s���Ă��邩�H

**�������@:**
```bash
# �Ǘ��҂Ƃ��Ď��s
# PowerShell���Ǘ��҂Ƃ��ĊJ��
cd path\to\project
python main.py
```

### ���: �^�C�}�[�G���[���o��

**�m�F����:**
1. �}���`�X���b�h�Ń^�C�}�[�𑀍삵�Ă��Ȃ����H
2. closeEvent ���Ă΂�Ă��邩�H

**�������@:**
```python
# �^�C�}�[����͕K�����C���X���b�h��
# �V�O�i��/�X���b�g�o�R�ŌĂяo��
```

---

## ? �֘A���\�[�X

### Qt �X���b�h�Z�[�t�e�B:
- [Qt Thread-Safety](https://doc.qt.io/qt-6/threads-qobject.html)
- [PyQt6 Signals and Slots](https://www.riverbankcomputing.com/static/Docs/PyQt6/)

### Python�t�@�C�����b�N:
- [pathlib.Path.unlink()](https://docs.python.org/3/library/pathlib.html#pathlib.Path.unlink)
- [Windows File Locking](https://docs.microsoft.com/en-us/windows/win32/fileio/locking-and-unlocking-byte-ranges-in-files)

---

## ? �`�F�b�N���X�g

�C��������������m�F:

- [x] `utils.py` - �t�@�C�����b�N�������P
- [x] `gui.py` - �^�C�}�[�Ǘ����P
- [x] `gui.py` - �X���b�h�N���[���A�b�v���P
- [x] `gui.py` - closeEvent�ǉ�
- [x] `winmergexlsx.py` - sys.exit�폜
- [x] �\���`�F�b�N����
- [ ] ���@�e�X�g�iExcel�t�@�C���J������ԁj
- [ ] �X�g���X�e�X�g�i�J��Ԃ����s�j
- [ ] �N���[���A�b�v�e�X�g�i�����I���j

---

**Status**: ? �C������
**Impact**: ? Critical (�A�v���P�[�V�����t���[�Y������)
**Test Required**: Yes (���@�e�X�g����)
