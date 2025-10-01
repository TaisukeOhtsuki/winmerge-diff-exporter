# WinMerge Diff to Excel Exporter

�t�H���_�Ԃ̍�����**WinMerge**�Ŕ�r���A���ʂ����₷��**Excel�t�@�C��**�Ƃ��ďo�͂���GUI�A�v���P�[�V�����ł��B

<img width="1002" height="522" alt="image" src="https:/## �Z�p�d�l

### �A�[�L�e�N�`��
- **GUI�t���[�����[�N**: PyQt6 6.9.1
- **Excel����**: openpyxl 3.1.5 (����Python�ACOM�s�v)
- **HTML���**: BeautifulSoup4 4.12.3 + lxml 5.3.0
- **������r**: WinMerge (�O���v���Z�X�AHTML�o�͌`��)
- **�}���`�X���b�h**: QThread�g�p��UI�u���b�N���
- **�G���[�n���h�����O**: 3�i�K�̃t�@�C���ۑ��헪

### ��Ȑ݌v�p�^�[��
- **MVC����**: UI�A�r�W�l�X���W�b�N�A�f�[�^�����𕪗�
- **�V�O�i��/�X���b�g**: Qt�񓯊��ʐM�p�^�[��
- **���[�J�[�X���b�h**: �����ԏ������o�b�N�O���E���h�Ŏ��s
- **���g���C���J�j�Y��**: �t�@�C�����b�N���̎����Ď��s

### �t�@�C�������o�A���S���Y��
WinMerge��HTML�o�̓t�@�C�����i`folder1_folder2_..._filename`�`���j����A�m���Ƀt�@�C�����𒊏o�F
- �t�H���_�L�[���[�h���o�i`modules`, `ctrl`, `tool`, `gui`, `etc` �Ȃǁj
- �g���q����t���X�L����
- �t�H���_�L�[���[�h�̎�����t�@�C�����Ƃ��ĔF��
- ��: `sc2stb_ctrl_v4_LINK_CTRLA_SP3_2.l` �� `CTRLA_SP3_2.l`

### �������o���W�b�N
- WinMerge�̍����F `FFEFCB05` (���F) �����o
- �R���e�L�X�g�s��: �O��4�s
- �אڂ��鍷���u���b�N�������}�[�W
- �t�H���_�s�̓����N�L���Ŏ��ʂ��ď��Or-attachments/assets/396c3c7a-3868-4199-898b-61229ab489f9" />

## �ڎ�

- [��ȋ@�\](#��ȋ@�\)
- [�V�X�e���v��](#�V�X�e���v��)
- [�C���X�g�[�����@](#�C���X�g�[�����@)
- [�g�p���@](#�g�p���@)
- [�v���W�F�N�g�\��](#�v���W�F�N�g�\��)
- [�Z�p�d�l](#�Z�p�d�l)
- [�ݒ�̃J�X�^�}�C�Y](#�ݒ�̃J�X�^�}�C�Y)
- [�g���u���V���[�e�B���O](#�g���u���V���[�e�B���O)
- [�J��](#�J��)
- [�ύX����](#�ύX����)
- [���C�Z���X](#���C�Z���X)

## ��ȋ@�\

- **�t�H���_�Ԃ̍�����r**: 2�̃t�H���_���r���A�ǉ��E�ύX�E�폜���ꂽ�t�@�C�����������o
- **Excel�`���ŏo��**: ��r���ʂ����₷��Excel�t�@�C���Ƃ��ĕۑ�
- **�ڍׂȍ����\��**: �t�@�C�����e�̍s���x���ł̍�����F�������ĕ\��
- **�h���b�O&�h���b�v�Ή�**: �t�H���_��GUI�ɒ��ڃh���b�O���ĊȒP�I��
- **�v���O���X�o�[**: �����i�s�󋵂����A���^�C���ŕ\��
- **�����V�[�g����**: 
  - **compare�V�[�g**: �����u���b�N�݂̂𒊏o�����ڍו\���i�t�@�C�������x���t���j
  - **�ʃt�@�C���V�[�g**: �e�t�@�C���̊��S�ȍ����i�t�@�C�������V�[�g���j
  - **Summary�V�[�g**: �t�@�C���ꗗ�ƕύX�󋵂̊T�v�i�t�H���_�s���������O�j

## �V�X�e���v��

- **Python 3.8�ȏ�** (Python 3.13.3�œ���m�F�ς�)
- **WinMerge** (�f�t�H���g�p�X: `C:\Program Files\WinMerge\WinMergeU.exe`)

### ? �o�[�W���� 2.0 �̎�ȉ��P�_

**Microsoft Excel �̃C���X�g�[���͕s�v�ɂȂ�܂����I**

- ? **Excel�Ȃ��œ���**: Excel���C���X�g�[������Ă��Ȃ��Ă����S�ɓ���
- ? **�t�@�C�����b�N�΍�**: Excel�Ńt�@�C�����J���Ă��Ă����s�\
  - 3�i�K�̕ۑ��헪�i���ڕۑ����ꎞ�t�@�C���o�R���^�C���X�^���v�t���j
- ? **�����ň���**: COM�ˑ���r�����A������Python���C�u�����ŏ���
- ? **���P���ꂽ�̍�**: 
  - �c�r���݂̂ŃX�b�L�����������ځi���r���Ȃ��j
  - �s�ԍ���́u.�v�������폜
  - ����������s�̂ݔw�i�F��\��
- ? **���m�ȃt�@�C�����\��**:
  - �V�[�g�����t�@�C�����Ɂi�p�X�`���ł͂Ȃ��j
  - compare�V�[�g�ł��������t�@�C������\��
  - �����p�X���ł��m���Ƀt�@�C�����𒊏o
- ? **�K�؂ȍ������o**:
  - WinMerge�̉��F�����s�𐳊m�Ɍ��o
  - Summary�V�[�g����t�H���_�s���������O

## �C���X�g�[�����@

### 1. ���|�W�g���̃N���[��
```bash
git clone https://github.com/TaisukeOhtsuki/winmerge-diff-exporter.git
cd winmerge-diff-exporter
```

### 2. ���z���̍쐬�i�����j
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/macOS
source venv/bin/activate
```

### 3. �ˑ��֌W�̃C���X�g�[��
```bash
pip install -r requirements.txt
```

**�K�v�ȃ��C�u����:**
- PyQt6 6.9.1 - GUI �t���[�����[�N
- openpyxl 3.1.5 - Excel �t�@�C������
- beautifulsoup4 4.12.3 - HTML �p�[�X
- lxml 5.3.0 - XML/HTML �p�[�T�[

### 4. WinMerge�̃C���X�g�[��
WinMerge�����C���X�g�[���̏ꍇ�́A[�����T�C�g](https://winmerge.org/)����_�E�����[�h���ăC���X�g�[�����Ă��������B

**�f�t�H���g�C���X�g�[���p�X:**
```
C:\Program Files\WinMerge\WinMergeU.exe
```

�J�X�^���p�X�̏ꍇ�� `src/core/config.py` �Őݒ��ύX�ł��܂��B

## �g�p���@

### 1. �A�v���P�[�V�����̋N��
```bash
python main.py
```

### 2. �t�H���_�̑I��
- **Base Folder**: ��r���̃t�H���_��I��
- **Comparison Target Folder**: ��r��̃t�H���_��I��  
- **Output File**: �o�͂���Excel�t�@�C�������w��

### 3. ���s
�uRun (Compare and Export to Excel)�v�{�^�����N���b�N���Ĕ�r���J�n���܂��B

### 4. ���ʂ̊m�F
�w�肵��Excel�t�@�C������������A�ȉ��̃V�[�g���܂܂�܂��F
- **compare�V�[�g**: �����u���b�N�݂̂𒊏o�i�R���e�L�X�g4�s�t���A�t�@�C�������x���\���j
- **Summary�V�[�g**: �t�@�C���ꗗ�ƕύX�󋵂̊T�v�i�t�H���_�s�͏��O�j
- **�ʃt�@�C���V�[�g**: �e�t�@�C���̊��S�ȍ����i�V�[�g���̓t�@�C�����A�s�ԍ��E�F�t���j

### 5. �t�@�C�����b�N���̓���
�o�̓t�@�C�����J����Ă���ꍇ�A�ȉ��̐헪�ŕۑ������݂܂��F
1. **���ڕۑ�** (5�񃊃g���C�A�i�K�I�ҋ@)
2. **�ꎞ�t�@�C���o�R** (�A�g�~�b�N�Ȗ��O�ύX)
3. **�^�C���X�^���v�t��** (��: `output_20250930_143052.xlsx`)

## �v���W�F�N�g�\��

```
winmerge-diff-exporter/
������ main.py                 # �A�v���P�[�V�����G���g���[�|�C���g
������ requirements.txt        # Python�ˑ��p�b�P�[�W
������ qt.conf                # Qt�ݒ�
������ LICENSE                # MIT���C�Z���X
������ README.md              # ���̃t�@�C��
��
������ src/                   # �\�[�X�R�[�h
��   ������ core/              # �R�A�r�W�l�X���W�b�N
��   ��   ������ common.py      # ���K�[�ƃ^�C�}�[
��   ��   ������ config.py      # �ݒ�Ǘ�
��   ��   ������ exceptions.py  # �J�X�^����O
��   ��   ������ utils.py       # �t�@�C�������Excel���`
��   ��   ������ winmergexlsx.py           # WinMerge����
��   ��   ������ diffdetailsheetcreater.py # �����ڍ׃V�[�g�쐬
��   ��
��   ������ converters/        # �t�@�C���ϊ�
��   ��   ������ html_to_excel.py  # HTML��Excel�ϊ��iCOM�s�v�j
��   ��
��   ������ ui/                # ���[�U�[�C���^�[�t�F�[�X
��       ������ gui.py         # PyQt6 GUI
��
������ docs/                  # �h�L�������g
��   ������ PROJECT_STRUCTURE.md       # �v���W�F�N�g�\���ڍ�
��   ������ EXCEL_COM_REMOVAL.md       # Excel COM�폜�̋Z�p�ڍ�
��   ������ FILE_LOCK_COMPLETE_FIX.md  # �t�@�C�����b�N�΍�
��   ������ REFACTORING_SUMMARY.md     # ���t�@�N�^�����O����
��   ������ RELEASE_NOTES_v2.0.md      # �����[�X�m�[�g
��
������ tests/                 # ���j�b�g�e�X�g
������ output/                # �o�̓t�@�C��
������ venv/                  # Python���z��
```

�ڍׂ� [`docs/PROJECT_STRUCTURE.md`](docs/PROJECT_STRUCTURE.md) ���Q�Ƃ��Ă��������B

## �Z�p�d�l

### �A�[�L�e�N�`��
- **GUI�t���[�����[�N**: PyQt6 6.9.1
- **Excel����**: openpyxl 3.1.5 (����Python�ACOM�s�v)
- **HTML���**: BeautifulSoup4 4.12.3 + lxml 5.3.0
- **������r**: WinMerge (�O���v���Z�X�AHTML�o�͌`��)
- **�}���`�X���b�h**: QThread�g�p��UI�u���b�N���
- **�G���[�n���h�����O**: 3�i�K�̃t�@�C���ۑ��헪

### ��Ȑ݌v�p�^�[��
- **MVC����**: UI�A�r�W�l�X���W�b�N�A�f�[�^�����𕪗�
- **�V�O�i��/�X���b�g**: Qt�񓯊��ʐM�p�^�[��
- **���[�J�[�X���b�h**: �����ԏ������o�b�N�O���E���h�Ŏ��s
- **���g���C���J�j�Y��**: �t�@�C�����b�N���̎����Ď��s


## �ݒ�̃J�X�^�}�C�Y

`src/core/config.py` �ňȉ��̐ݒ��ύX�ł��܂��F

- **WinMerge�p�X**: �J�X�^���C���X�g�[������w��
  ```python
  executable_path: str = r'C:\Program Files\WinMerge\WinMergeU.exe'
  ```
- **�����F**: WinMerge�����s�̐F�R�[�h
  ```python
  yellow_color: str = 'FFEFCB05'  # WinMerge diff color (yellow)
  ```
- **�R���e�L�X�g�s��**: �����O��̕\���s��
  ```python
  context_lines: int = 4
  ```
- **��**: Excel��̕��ݒ�i`diff_formats` �����j
- **�t�H���_�L�[���[�h**: �t�@�C�������o���ɔF������t�H���_���i`_extract_filename_from_stem` ���\�b�h�j

## �g���u���V���[�e�B���O

### WinMerge��������Ȃ�
```
WinMergeNotFoundError: WinMerge not found at ...
```
�� `src/core/config.py` �� `winmerge_path` �𐳂����p�X�ɐݒ肵�Ă��������B

### �t�@�C�����ۑ��ł��Ȃ�
�o�̓t�@�C�����J����Ă���ꍇ�A�^�C���X�^���v�t���t�@�C�����쐬����܂��B
��: `output_20250930_143052.xlsx`

### DPI�X�P�[�����O���
��DPI����GUI���������\������Ȃ��ꍇ�A`main.py` �� DPI�ݒ肪�����������܂��B

## �J��

### �V�@�\�̒ǉ�
1. �R�A���W�b�N �� `src/core/`
2. UI �R���|�[�l���g �� `src/ui/`
3. �t�@�C���ϊ� �� `src/converters/`
4. �e�X�g �� `tests/`
5. �h�L�������g �� `docs/`

### �R�[�h�X�^�C��
- ���΃C���|�[�g: �p�b�P�[�W�� (`.module`)
- ��΃C���|�[�g: �p�b�P�[�W�O (`src.package.module`)
- UTF-8�G���R�[�f�B���O
- �p��R�����g����

### �e�X�g���s
```bash
# �\���`�F�b�N
python -m py_compile main.py

# �A�v���P�[�V�������s
python main.py
```

## �ύX����

### Version 2.1 (2025-10-01)
- ? �V�[�g�����t�@�C�����ɕύX�i�p�X�`��������P�j
- ? compare�V�[�g�̃t�@�C�����\�����C��
- ? ���m�ȃt�@�C�������o�A���S���Y������
- ? �����F�̏C���i�D�F�����F `FFEFCB05`�j
- ? Summary�V�[�g����t�H���_�s�����O
- ? ���ۂ̕ۑ��t�@�C������GUI�ɕ\��
- ? �f�o�b�O���O�̒ǉ�

### Version 2.0 (2025-09)
- ? Excel COM�ˑ������S�폜
- ? �t�@�C�����b�N�΍�̎���
- ? �v���W�F�N�g�\���̍ĕҐ�
- ? �̍ق̉��P�i�c�r���̂݁A�s�ԍ��́u.�v�폜�j
- ? DPI�X�P�[�����O�Ή�
- ? �A�v���P�[�V�����t���[�Y���̏C��

�ڍׂ� [`docs/RELEASE_NOTES_v2.0.md`](docs/RELEASE_NOTES_v2.0.md) ���Q�Ƃ��Ă��������B

## ���C�Z���X

MIT License - �ڍׂ�[LICENSE](LICENSE)�t�@�C�����Q�Ƃ��Ă��������B

## �v��

�o�O�񍐂�@�\��Ă�[Issues](https://github.com/TaisukeOhtsuki/winmerge-diff-exporter/issues)�ł��肢���܂��B

�v�����N�G�X�g�����}���܂��I

## �ӎ�

���̃\�t�g�E�F�A��[winmerge_xlsx](https://github.com/y-tetsu/winmerge_xlsx.git)�̃R�[�h���x�[�X�ɊJ������܂����B

## ���

**TaisukeOhtsuki** - [GitHub](https://github.com/TaisukeOhtsuki)

---

? ���̃v���W�F�N�g�����ɗ�������A�X�^�[�����肢���܂��I
