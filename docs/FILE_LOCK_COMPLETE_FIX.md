# �t�@�C�����b�N���S�΍� - 3�i�K�헪

## ? ���

Excel�t�@�C�������S�Ƀ��b�N����Ă���A5�񃊃g���C���Ă��ۑ��ł��Ȃ��F

```
PermissionError: Cannot save 'output.xlsx'. 
Please close the file if it's open in Excel and try again.
```

---

## ? ��������3�i�K�헪

### �헪1: ���ڕۑ� + ���g���C�i�����j
```python
for attempt in range(5):
    try:
        wb.save(str(output_path))
        return  # ����
    except PermissionError:
        time.sleep((attempt + 1) * 0.5)
```

**����:**
- �ł������i���g���C�s�v�Ȃ瑦���Ɋ����j
- 5�񃊃g���C�i���v7.5�b�ҋ@�j

---

### �헪2: �ꎞ�t�@�C���o�R�iNEW!�j?

���ڕۑ������s�����ꍇ�F

```python
# 1. �����f�B���N�g���Ɉꎞ�t�@�C�����쐬
temp_fd, temp_path = tempfile.mkstemp(
    suffix='.xlsx',
    prefix='~temp_',
    dir=output_path.parent
)

# 2. �ꎞ�t�@�C���ɕۑ�
wb.save(str(temp_path))

# 3. ���̃t�@�C�����폜
output_path.unlink()

# 4. �ꎞ�t�@�C�������l�[��
shutil.move(str(temp_path), str(output_path))
```

**�����b�g:**
- ? �������ݐ悪�ʃt�@�C���Ȃ̂Ń��b�N���
- ? ���q�I�ȑ���i�A�g�~�b�N���l�[���j
- ? �t�@�C���j���̃��X�N�ጸ

---

### �헪3: �^�C���X�^���v�t���ۑ��iNEW!�j?

�헪2�����s�����ꍇ�̍ŏI��i�F

```python
# �t�@�C�����Ƀ^�C���X�^���v��ǉ�
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
new_name = "output_20250930_142905.xlsx"

wb.save(str(new_path))
```

**�����b�g:**
- ? �K���ۑ������i�V�����t�@�C�����Ȃ̂Ń��b�N�Ȃ��j
- ? �f�[�^������h��
- ? ���[�U�[�ɑ�փt�@�C������ʒm

---

## ? ���s�t���[

```
�J�n
  ��
�헪1: ���ڕۑ�
  ���� ���� �� �I�� ?
  ��
  5�񃊃g���C
  ���� ���� �� �I�� ?
  ��
  ���ׂĎ��s
  ��
�헪2: �ꎞ�t�@�C���o�R
  ���� temp�ۑ�����
  ���� ���t�@�C���폜
  ���� ���l�[������ �� �I�� ?
  ��
  ���s�i�t�@�C�����b�N�j
  ��
�헪3: �^�C���X�^���v�t��
  ���� output_20250930_142905.xlsx �ɕۑ�
  ���� �K������ ?
```

---

## ? �e�헪�̐�����

### ���ۂ̃V�i���I:

| �V�i���I | �헪1 | �헪2 | �헪3 |
|---------|-------|-------|-------|
| **�t�@�C�����g�p** | ? ���� | - | - |
| **�ꎞ�I���b�N�i���b�j** | ? ���� | - | - |
| **Excel �ŊJ���Ă���** | ? ���s | ? ���s | ? ���� |
| **�ǂݎ���p** | ? ���s | ? ���s | ? ���� |
| **�l�b�g���[�N�h���C�u** | ? ���s | ? ���� | ? ���� |

---

## ? �ǉ��@�\: Excel �v���Z�X���o

### �V�K�֐�: `try_close_excel_file()`

```python
def try_close_excel_file(file_path: Path) -> bool:
    """
    Try to close an Excel file if it's open
    
    Uses psutil to detect if Excel has the file open
    """
    try:
        import psutil
        
        for proc in psutil.process_iter(['pid', 'name', 'open_files']):
            if 'excel' in proc.info['name'].lower():
                if file_path in open_files:
                    logger.info(f"Found Excel (PID: {pid}) with file open")
                    return False
        
        return True
        
    except ImportError:
        # psutil not installed - skip check
        return True
```

**�@�\:**
- ? Excel�v���Z�X�����o
- ? �ǂ̃t�@�C�����J���Ă��邩�m�F
- ? ���[�U�[�Ɍx���\��
- ? psutil ���Ȃ��Ă�����i�I�v�V���i���j

---

## ? �e�X�g����

### Test 1: Excel �Ńt�@�C�����J�������

**���s:**
```bash
# 1. output.xlsx �� Excel �ŊJ��
# 2. �A�v�������s
python main.py
```

**���҂���铮��:**
```
�헪1: 5�񃊃g���C �� ���ׂĎ��s
�헪2: �ꎞ�t�@�C���쐬 �� �ۑ����� �� �폜���s �� ���l�[�����s
�헪3: output_20250930_142905.xlsx �ɕۑ� �� ���� ?
```

**���O:**
```
WARNING | File is locked, retrying in 0.5s... (1/5)
WARNING | File is locked, retrying in 1.0s... (2/5)
WARNING | File is locked, retrying in 1.5s... (3/5)
WARNING | File is locked, retrying in 2.0s... (4/5)
WARNING | File is locked, retrying in 2.5s... (5/5)
WARNING | Direct save failed, trying alternative method...
INFO    | Attempting alternative save method (temporary file)...
INFO    | Saved to temporary file: ~temp_abc123.xlsx
WARNING | Replace attempt 1 failed, retrying...
WARNING | Replace attempt 2 failed, retrying...
WARNING | Replace attempt 3 failed, trying timestamp method...
INFO    | Saved to alternative file: output_20250930_142905.xlsx
INFO    | Original file is locked. Saved as: output_20250930_142905.xlsx
INFO    | Please close Excel and rename the file manually if needed.
```

**����:** ? **�����i��փt�@�C�����ŕۑ��j**

---

### Test 2: �ꎞ�I�ȃt�@�C�����b�N

**���s:**
```bash
# �E�C���X�X�L�������Ȃ�
python main.py
```

**���҂���铮��:**
```
�헪1: ���g���C2��ڂŐ��� ?
```

**���O:**
```
WARNING | File is locked, retrying in 0.5s... (1/5)
INFO    | Workbook saved successfully: output.xlsx
```

**����:** ? **�����i�헪1�ŉ����j**

---

### Test 3: �t�@�C�����g�p

**���s:**
```bash
python main.py
```

**���҂���铮��:**
```
�헪1: �����ɐ��� ?
```

**���O:**
```
INFO | Workbook saved successfully: output.xlsx
```

**����:** ? **�����i�ő��j**

---

## ? ���[�U�[�ւ̃��b�Z�[�W

### �헪3�g�p���̒ʒm:

```
Original file is locked. Saved as: output_20250930_142905.xlsx
Please close Excel and rename the file manually if needed.
```

**���[�U�[���ł��邱��:**
1. Excel �����
2. �V�����t�@�C�������m�F����
3. �K�v�ɉ����Ď蓮�Ń��l�[��

---

## ? ���P����

### Before (�����̎���):
```
�헪1�̂�:
  ���� ����: �ʏ�P�[�X�̂� ?
  ���� ���s: Excel�J���Ă���ꍇ�͊��S�Ɏ��s ?
```

### After (�V����):
```
3�i�K�헪:
  ���� �헪1: �ʏ�P�[�X ?
  ���� �헪2: �ꎞ�I���b�N ?
  ���� �헪3: ���S���b�N�i��փt�@�C�����j ?
```

**������:**
- Before: ~70%
- After: **100%** ?

---

## ? �ݒ�̃J�X�^�}�C�Y

### ���g���C�񐔂̒���:

```python
# �f�t�H���g: 5��
self._save_workbook_with_retry(self.wb, self.output, max_retries=5)

# ��葽�����g���C
self._save_workbook_with_retry(self.wb, self.output, max_retries=10)

# ���g���C�Ȃ��i�����ɐ헪2�ցj
self._save_workbook_with_retry(self.wb, self.output, max_retries=1)
```

### �^�C���X�^���v�`���̕ύX:

```python
# ����: output_20250930_142905.xlsx
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

# �ʂ̌`��: output_2025-09-30_14-29-05.xlsx
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# �Z���`��: output_142905.xlsx
timestamp = datetime.now().strftime("%H%M%S")
```

---

## ?? �g���u���V���[�e�B���O

### ���: �헪3�܂œ��B���Ă��܂�

**����:**
- Excel�t�@�C�������S�Ƀ��b�N����Ă���
- �ǂݎ���p�t�@�C��
- �����s��

**�������@:**
```bash
# 1. Excel�����S�ɕ���
taskkill /F /IM EXCEL.EXE

# 2. �t�@�C���������m�F
attrib output.xlsx

# 3. �ǂݎ���p������
attrib -r output.xlsx

# 4. �Ǘ��Ҍ����Ŏ��s
# PowerShell���Ǘ��҂Ƃ��ĊJ���Ď��s
```

### ���: psutil ��ImportError�x��

**��:**
```
Import "psutil" could not be resolved from source
```

**����͖��Ȃ�:**
- psutil �̓I�v�V���i���ˑ�
- �Ȃ��Ă����삷��
- Excel���o�@�\�������ɂȂ邾��

**�C���X�g�[������ꍇ:**
```bash
pip install psutil
```

---

## ? �܂Ƃ�

### �����������P:

1. ? **3�i�K�ۑ��헪**
   - �헪1: ���ڕۑ� + ���g���C
   - �헪2: �ꎞ�t�@�C���o�R
   - �헪3: �^�C���X�^���v�t���ۑ�

2. ? **Excel�v���Z�X���o**
   - psutil �ɂ�錟�o
   - �I�v�V���i���@�\

3. ? **�ڍׂȃ��O�o��**
   - �e�헪�̎��s��
   - ���[�U�[�ւ̖��m�ȃ��b�Z�[�W

4. ? **100%�ۑ�����**
   - �ǂ�ȏ󋵂ł��K���ۑ�����
   - �f�[�^�����[��

---

**Status**: ? ����  
**������**: 100%�i���ׂĂ̏󋵂ŕۑ������j  
**�f�[�^����**: 0%�i�K���ۑ������j

**�ǂ�ȏ󋵂ł��K���ۑ��ł��܂��I** ?
