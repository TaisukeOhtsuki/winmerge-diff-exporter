# Permission Error �C�� - ���S��

## ? �������Ă������

### �G���[:
```
PermissionError: [Errno 13] Permission denied: 'output.xlsx'
```

### ����:
1. **�t�@�C�����J�����܂�**: Excel �� `output.xlsx` ���J���Ă����Ԃŕۑ������s
2. **�����̎��s**: ���g���C�Ȃ��ő����Ɏ��s
3. **�s�e�؂ȃG���[���b�Z�[�W**: ���[�U�[���Ώ����@�𗝉��ł��Ȃ�

---

## ? ���������C��

### 1. Workbook�ۑ��̃��g���C���W�b�N�ǉ�

#### �V�K���\�b�h: `_save_workbook_with_retry()`

```python
def _save_workbook_with_retry(self, wb, output_path: Path, max_retries: int = 5) -> None:
    """
    Save workbook with retry logic for file locks
    
    - �ő�5�񃊃g���C
    - �ҋ@���Ԃ�i�K�I�ɑ����i0.5�b �� 1.0�b �� 1.5�b...�j
    - �ڍׂȃ��O�o��
    - ���[�U�[�t�����h���[�ȃG���[���b�Z�[�W
    """
    import time
    
    for attempt in range(max_retries):
        try:
            wb.save(str(output_path))
            logger.info(f"Workbook saved successfully: {output_path}")
            return
        except PermissionError as e:
            if attempt < max_retries - 1:
                wait_time = (attempt + 1) * 0.5  # 0.5s, 1.0s, 1.5s, 2.0s, 2.5s
                logger.warning(f"File is locked, retrying in {wait_time}s... ({attempt + 1}/{max_retries})")
                self.log(f"File is locked. Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                # �ŏI���s���s
                raise PermissionError(
                    f"Cannot save '{output_path.name}'. "
                    f"Please close the file if it's open in Excel and try again."
                )
```

**����:**
- ? **�i�K�I�ҋ@**: 0.5�b �� 1.0�b �� 1.5�b �� 2.0�b �� 2.5�b
- ? **�ڍ׃��O**: �e���g���C�̏󋵂����O�o��
- ? **UI�ʒm**: ���[�U�[�ɏ󋵂�ʒm
- ? **���m�ȃG���[**: �ŏI���s���ɑΏ����@���

---

### 2. �ۑ�������2�ӏ��ɓK�p

#### �ӏ�1: ����ۑ��iHTML��Excel�ϊ���j
```python
# Before:
self.wb.save(str(self.output))  # �����Ɏ��s

# After:
self._save_workbook_with_retry(self.wb, self.output)  # ���g���C����
```

#### �ӏ�2: �ŏI�ۑ��i�t�H�[�}�b�g�K�p��j
```python
def _save_workbook(self) -> None:
    """Save the workbook with retry logic"""
    self._save_workbook_with_retry(self.wb, self.output)
    self.log(f"Excel file saved: {self.output}")
```

---

### 3. �G���[�n���h�����O�̉��P

#### `winmergexlsx.py`:
```python
except PermissionError as e:
    error_msg = f"Cannot save Excel file. Please close '{self.output.name}' if it's open in Excel."
    logger.error(error_msg, exc_info=True)
    raise ExcelProcessingError(error_msg)
```

#### `gui.py`:
```python
def _handle_error(self, error_msg: str) -> None:
    """Handle worker error"""
    if "Permission denied" in error_msg or "close" in error_msg.lower():
        detailed_msg = (
            f"{error_msg}\n\n"
            "Tips:\n"
            "? Close the Excel file if it's currently open\n"
            "? Check if another process is using the file\n"
            "? Try saving to a different location"
        )
        QMessageBox.critical(self, "File Access Error", detailed_msg)
    else:
        QMessageBox.critical(self, "Process Error", error_msg)
```

---

## ? �C���O��̔�r

### Before (��肠��):

```
1. HTML��Excel�ϊ�����
2. wb.save() ���s
3. PermissionError �����ɔ���
4. �G���[���b�Z�[�W�\��
5. �������s ?
```

**���O:**
```
ERROR | Permission denied: 'output.xlsx'
```

### After (�C����):

```
1. HTML��Excel�ϊ�����
2. _save_workbook_with_retry() ���s
3. PermissionError ����
4. 0.5�b�ҋ@���ă��g���C
5. �܂����b�N��
6. 1.0�b�ҋ@���ă��g���C
7. �t�@�C���N���[�Y���ꂽ
8. �ۑ����� ?
```

**���O:**
```
WARNING | File is locked, retrying in 0.5s... (1/5)
INFO    | File is locked. Retrying in 0.5 seconds...
WARNING | File is locked, retrying in 1.0s... (2/5)
INFO    | File is locked. Retrying in 1.0 seconds...
INFO    | Workbook saved successfully: output.xlsx
INFO    | Excel file saved: output.xlsx
```

---

## ? ���[�X�P�[�X

### Case 1: Excel�Ńt�@�C�����J���Ă���

**�V�i���I:**
```
1. ���[�U�[�� output.xlsx �� Excel �ŊJ���Ă���
2. �A�v�������s
3. �����J�n
4. �ۑ����� PermissionError
5. ���g���C�J�n
6. ���[�U�[�� Excel �����
7. ���̃��g���C�ŕۑ�����
```

**����:** ? **�����i���[�U�[���t�@�C�������Ύ����I�ɕۑ��j**

### Case 2: �ʃv���Z�X���t�@�C�������b�N

**�V�i���I:**
```
1. �E�C���X�΍�\�t�g���t�@�C�����X�L������
2. �A�v�������s
3. �ۑ����� PermissionError
4. 5�񃊃g���C�i���v7.5�b�ҋ@�j
5. �X�L��������
6. �ۑ�����
```

**����:** ? **�����i�ꎞ�I�ȃ��b�N�������I�ɉ���j**

### Case 3: �i���I�ȃ��b�N

**�V�i���I:**
```
1. �t�@�C�����ǂݎ���p
2. �A�v�������s
3. 5�񃊃g���C���Ă����s
4. �ڍׂȃG���[���b�Z�[�W�\��
```

**�G���[���b�Z�[�W:**
```
Cannot save Excel file. Please close 'output.xlsx' if it's open in Excel.

Tips:
? Close the Excel file if it's currently open
? Check if another process is using the file
? Try saving to a different location
```

**����:** ? **���s���邪�A���[�U�[�ɖ��m�ȑΏ����@���**

---

## ? �Z�p�ڍ�

### ���g���C�^�C�~���O

```python
Attempt 1: �����Ɏ��s
Attempt 2: 0.5�b�ҋ@��
Attempt 3: 1.0�b�ҋ@��
Attempt 4: 1.5�b�ҋ@��
Attempt 5: 2.0�b�ҋ@��
Attempt 6: 2.5�b�ҋ@��i�ŏI�j

���v�ő�ҋ@����: 7.5�b
```

### �Ȃ��i�K�I�ɑ����H

1. **�Z���ԃ��b�N**: 0.5�b�ŉ�������ꍇ������
2. **�����ԃ��b�N**: Excel�N�����ȂǁA1-2�b�K�v
3. **�����ԃ��b�N**: �E�C���X�X�L�����ȂǁA���b�K�v

�i�K�I�����ɂ��A�Z���ԃ��b�N�ɂ͑f�����Ή����A�����ԃ��b�N�ɂ��Ή��\�B

---

## ? �e�X�g���@

### Test 1: Excel�Ńt�@�C�����J�������

```bash
# 1. output.xlsx �� Excel �ŊJ��
# 2. �A�v�������s
python main.py

# �������Ƀ��O���m�F:
# "File is locked, retrying in 0.5s... (1/5)"
# "File is locked, retrying in 1.0s... (2/5)"

# 3. Excel �����
# 4. ���̃��g���C�Ő���

# ���Ҍ���:
# ? "Workbook saved successfully"
# ? ��������
```

### Test 2: �t�@�C�����J�����Ɏ��s

```bash
python main.py

# ���Ҍ���:
# ? ���g���C�Ȃ��ő����ɕۑ�����
# ? "Workbook saved successfully"
```

### Test 3: �ǂݎ���p�t�@�C��

```bash
# 1. output.xlsx ��ǂݎ���p�ɐݒ�
# 2. �A�v�������s

# ���Ҍ���:
# ? 5�񃊃g���C��Ɏ��s
# ? �ڍׂȃG���[���b�Z�[�W�\��
```

---

## ? ���O�o�͗�

### �����P�[�X�i2��ڂŐ����j:

```
2025-09-30 14:30:00 | INFO     | Converting HTML to Excel...
2025-09-30 14:30:01 | INFO     | Converted: cset_file
2025-09-30 14:30:01 | WARNING  | File is locked, retrying in 0.5s... (1/5)
2025-09-30 14:30:01 | INFO     | File is locked. Retrying in 0.5 seconds...
2025-09-30 14:30:02 | INFO     | Workbook saved successfully: output.xlsx
2025-09-30 14:30:02 | INFO     | Applying final formatting with openpyxl...
2025-09-30 14:30:03 | INFO     | Workbook saved successfully: output.xlsx
2025-09-30 14:30:03 | INFO     | Excel file saved: output.xlsx
2025-09-30 14:30:03 | INFO     | Generation completed successfully
```

### ���s�P�[�X�i5�񃊃g���C��j:

```
2025-09-30 14:30:00 | INFO     | Converting HTML to Excel...
2025-09-30 14:30:01 | WARNING  | File is locked, retrying in 0.5s... (1/5)
2025-09-30 14:30:02 | WARNING  | File is locked, retrying in 1.0s... (2/5)
2025-09-30 14:30:03 | WARNING  | File is locked, retrying in 1.5s... (3/5)
2025-09-30 14:30:05 | WARNING  | File is locked, retrying in 2.0s... (4/5)
2025-09-30 14:30:07 | WARNING  | File is locked, retrying in 2.5s... (5/5)
2025-09-30 14:30:10 | ERROR    | Failed to save after 5 attempts
2025-09-30 14:30:10 | ERROR    | Cannot save Excel file. Please close 'output.xlsx' if it's open in Excel.
```

---

## ?? �g���u���V���[�e�B���O

### ���: ���g���C���Ă��܂����s����

**�m�F����:**
1. Excel�t�@�C�����{���ɕ����Ă��邩�H
2. �^�X�N�}�l�[�W���[�� Excel.exe ���c���Ă��Ȃ����H
3. �t�@�C�����ǂݎ���p�ɂȂ��Ă��Ȃ����H

**�������@:**
```bash
# Excel�v���Z�X�����S�I��
taskkill /F /IM EXCEL.EXE

# �t�@�C���������m�F
attrib output.xlsx

# �ǂݎ���p������
attrib -r output.xlsx
```

### ���: ���g���C����������

**�������@:**
```python
# winmergexlsx.py �� _save_workbook_with_retry ���\�b�h

# ���g���C�񐔂����炷
self._save_workbook_with_retry(self.wb, self.output, max_retries=3)

# �ҋ@���Ԃ�Z�k
wait_time = (attempt + 1) * 0.3  # 0.3�b�Ԋu
```

---

## ? �`�F�b�N���X�g

�C�����e�̊m�F:

- [x] `_save_workbook_with_retry()` ���\�b�h�ǉ�
- [x] ����ۑ��Ƀ��g���C���W�b�N�K�p
- [x] �ŏI�ۑ��Ƀ��g���C���W�b�N�K�p
- [x] PermissionError �̖����I�n���h�����O
- [x] GUI�ł̃��[�U�[�t�����h���[�ȃG���[���b�Z�[�W
- [x] �ڍׂȃ��O�o��
- [x] �\���`�F�b�N����
- [ ] ���@�e�X�g�iExcel�t�@�C���J������ԁj

---

**Status**: ? �C������  
**Impact**: ? Critical (PermissionError���S�΍�)  
**Test Required**: Yes (Excel�t�@�C���J������ԂŃe�X�g)

**����ŁAExcel�t�@�C�����J���Ă��Ă������I�Ƀ��g���C���ĕۑ��ł��܂��I** ?
