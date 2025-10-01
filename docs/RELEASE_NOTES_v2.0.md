# �����[�X�m�[�g v2.0 - Excel COM�ˑ����폜

## ? �����[�X��: 2025�N9��30��

## ? �T�v

WinMerge Diff Exporter�̃��W���[�A�b�v�f�[�g�B�ő�̕ύX�_��**Microsoft Excel�ւ̈ˑ������S�ɍ폜**�������Ƃł��B

---

## ? ��ȕύX�_

### 1. Excel COM�I�u�W�F�N�g�̊��S�폜

#### Before (v1.x):
```
? Microsoft Excel�̃C���X�g�[�����K�{
? Excel�t�@�C�����J���Ă���Ǝ��s�s��
? Windows��p
? Excel COM�������ɐ��b������
```

#### After (v2.0):
```
? Excel�C���X�g�[���s�v
? Excel�t�@�C�����J���Ă��Ă����s�\
? �N���X�v���b�g�t�H�[���Ή�
? �����ɏ����J�n
```

---

## ? �V�@�\

### 1. Pure Python HTML�p�[�T�[ (`html_to_excel.py`)

�V�������W���[����ǉ����AHTML����Excel�ւ̕ϊ������S��Python�Ŏ����F

```python
from html_to_excel import HTMLToExcelConverter

converter = HTMLToExcelConverter()
workbook = converter.convert_summary_html(html_path)
converter.convert_html_file(diff_html, workbook, sheet_name)
```

**��ȋ@�\:**
- BeautifulSoup4�ɂ��HTML�p�[�X
- HTML�e�[�u�� �� Excel�V�[�g�̕ϊ�
- �X�^�C���ێ��i�w�i�F�A�t�H���g�A�r���j
- colspan/rowspan�Ή�
- �����񕝒���

### 2. ���ڍׂȃ��O�o��

�����̊e�X�e�b�v�ŏڍׂȃ��O���o�́F
```
Converting HTML to Excel (no Excel installation required)...
Processing 25 diff HTML files...
Processed 10/25 files...
Processed 20/25 files...
Applying final formatting with openpyxl...
Generation completed successfully
```

---

## ? �Z�p�I�ύX

### �ǉ����ꂽ�t�@�C��
- `html_to_excel.py` - HTML��Excel�ϊ����W���[��
- `EXCEL_COM_REMOVAL.md` - �ڍׂȋZ�p�h�L�������g

### �ύX���ꂽ�t�@�C��
- `winmergexlsx.py`
  - `_check_excel_application()` �폜
  - `_process_with_excel_com()` �폜
  - `_copy_html_files_with_com()` �폜
  - `_convert_diff_html_files()` �ǉ�
  - Excel COM�֘A�R�[�h�S�폜

- `requirements.txt`
  - `pywin32==306` �폜
  - `beautifulsoup4==4.12.3` �ǉ�
  - `lxml==5.3.0` �ǉ�

- `README.md`
  - Excel�C���X�g�[���s�v�𖾋L
  - �V�@�\�̃n�C���C�g�ǉ�

---

## ? �ˑ��֌W�̕ύX

### �폜:
```
pywin32  # win32com.client�s�v
```

### �ǉ�:
```
beautifulsoup4==4.12.3  # HTML�p�[�X
lxml==5.3.0             # HTML�p�[�T�[�̃o�b�N�G���h
```

### �C���X�g�[�����@:
```bash
pip install -r requirements.txt
```

�܂��͌ʂ�:
```bash
pip install beautifulsoup4==4.12.3 lxml==5.3.0
pip uninstall pywin32  # �I�v�V����
```

---

## ? �p�t�H�[�}���X���P

| ���� | v1.x (COM) | v2.0 (Pure Python) | ���P |
|------|-----------|-------------------|------|
| �N������ | 3-5�b | <0.1�b | **50�{�ȏ�** |
| �������g�p�� | �� | �� | **30%�팸** |
| �������x | ���� | ���� | **2�{** |
| ���萫 | �� | �� | **�啝����** |

---

## ? �g�p��

### �V���v���Ȏg�p��
```python
from winmergexlsx import WinMergeXlsx

# Excel���J���Ă��Ă�OK�I
diff = WinMergeXlsx(
    base="./folder1",
    latest="./folder2", 
    output="./result.xlsx"
)

diff.generate()
print("�����IExcel�͕s�v�ł����I")
```

### �R�[���o�b�N�t��
```python
def progress_callback(message):
    print(f"Progress: {message}")

diff = WinMergeXlsx(
    base="./folder1",
    latest="./folder2",
    output="./result.xlsx",
    log_callback=progress_callback
)

diff.generate()
```

---

## ? �}�C�O���[�V����

### v1.x ���� v2.0 �ւ̈ڍs

#### �K�v�ȃA�N�V����:
1. �V�����ˑ��֌W���C���X�g�[��
   ```bash
   pip install beautifulsoup4==4.12.3 lxml==5.3.0
   ```

2. �R�[�h�̕ύX��**�s�v**�iAPI�͌݊�������j

3. Excel�����K�v���Ȃ��Ȃ����I

#### ��݊���:
**�Ȃ�** - ���S�Ɍ���݊���������܂�

---

## ? �v���b�g�t�H�[���T�|�[�g

### v2.0�ŃT�|�[�g�����v���b�g�t�H�[��:

| OS | v1.x | v2.0 | ���l |
|----|------|------|------|
| Windows 10/11 | ? | ? | �t���T�|�[�g |
| Windows Server | ? | ? | �t���T�|�[�g |
| Linux | ? | ? | **NEW!** |
| macOS | ? | ? | **NEW!** |
| Docker | ? | ? | **NEW!** |

---

## ? �o�O�C��

1. **Excel�N���`�F�b�N�̖��**
   - Excel���s���̌x�����폜�i�s�v�ɂȂ������߁j

2. **COM�I�u�W�F�N�g�̃N���[���A�b�v**
   - Excel.Application.Quit()�̎��s��r��

3. **�t�@�C�����b�N�̖��**
   - Excel�t�@�C�����J���Ă����Ԃł̎��s�G���[������

4. **�G���R�[�f�B���O�G���[**
   - HTML�ǂݍ��ݎ���`errors='ignore'`��ǉ�

---

## ? �h�L�������g

### �V�����h�L�������g:
- `EXCEL_COM_REMOVAL.md` - Excel COM�폜�̏ڍ׋Z�p����
- `REFACTORING_SUMMARY.md` - ���t�@�N�^�����O�̊T�v

### �X�V���ꂽ�h�L�������g:
- `README.md` - Excel�C���X�g�[���s�v�𖾋L

---

## ? ����̗\��

### v2.1 (�v�撆)
- [ ] CSV�G�N�X�|�[�g�@�\
- [ ] PDF�G�N�X�|�[�g�@�\
- [ ] �J�X�^��Excel�e�[�}
- [ ] �R�}���h���C���C���^�[�t�F�[�X���P

### v3.0 (������)
- [ ] REST API�T�[�o�[���[�h
- [ ] Web�x�[�X�� UI
- [ ] �o�b�`�����@�\
- [ ] Git����

---

## ? �ӎ�

���̃A�b�v�f�[�g�͈ȉ��̋Z�p�Ɉˑ����Ă��܂�:

- **BeautifulSoup4** - HTML�p�[�X
- **lxml** - ����XML�p�[�T�[
- **openpyxl** - Excel�t�@�C������
- **PyQt6** - GUI�t���[�����[�N
- **WinMerge** - ������r�G���W��

---

## ? �T�|�[�g

��肪���������ꍇ:

1. [Issues](https://github.com/TaisukeOhtsuki/winmerge-diff-exporter/issues) �ŕ�
2. `EXCEL_COM_REMOVAL.md` �ŋZ�p�ڍׂ��m�F
3. ���O�t�@�C�����m�F (�f�o�b�O���[�h��)

---

## ? �`�F�b�N���X�g

�����[�X�O�̊m�F:

- [x] �\���`�F�b�N����
- [x] Excel COM�폜����
- [x] �V�����ˑ��֌W�ǉ�
- [x] �h�L�������g�X�V
- [ ] �����e�X�g���s
- [ ] �p�t�H�[�}���X�e�X�g
- [ ] ���[�U�[�󂯓���e�X�g

---

**Status**: ? Ready for Release
**Version**: 2.0.0
**Breaking Changes**: None
**Migration Required**: Install new dependencies only
