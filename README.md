# WinMerge Diff to Excel Exporter

�t�H���_�Ԃ̍�����**WinMerge**�Ŕ�r���A���ʂ����₷��**Excel�t�@�C��**�Ƃ��ďo�͂���GUI�A�v���P�[�V�����ł��B

<img width="1002" height="522" alt="image" src="https://github.com/user-attachments/assets/396c3c7a-3868-4199-898b-61229ab489f9" />

## �@�\

- **�t�H���_�Ԃ̍�����r**: 2�̃t�H���_���r���A�ǉ��E�ύX�E�폜���ꂽ�t�@�C�������o
- **Excel�t�@�C���o��**: ��r���ʂ�Excel�`���ŕۑ�
- **�ڍׂȍ����\��**: �t�@�C�����e�̍s���x���ł̍�����\��
- **�h���b�O&�h���b�v�Ή�**: �t�H���_��GUI�ɒ��ڃh���b�O���đI���\
- **�v���O���X�o�[**: �����i�s�󋵂����A���^�C���ŕ\��

## �K�v�Ȋ�

- **Windows 10/11** (WinMerge���K�v)
- **Python 3.8�ȏ�**
- **WinMerge** (�ȉ��̃p�X�ɃC���X�g�[������Ă���K�v������܂�)
  ```
  C:\Program Files\WinMerge\WinMergeU.exe
  ```

## �C���X�g�[�����@

### 1. ���|�W�g���̃N���[��
```bash
git clone https://github.com/your-username/winmerge-diff-exporter.git
cd winmerge-diff-exporter
```

### 2. ���z���̍쐬�i�����j
```bash
python -m venv venv
venv\Scripts\activate
```

### 3. �ˑ��֌W�̃C���X�g�[��
```bash
pip install -r requirements.txt
```

### 4. WinMerge�̃C���X�g�[��
WinMerge�����C���X�g�[���̏ꍇ�́A[�����T�C�g](https://winmerge.org/)����_�E�����[�h���ăC���X�g�[�����Ă��������B

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
- **compare**: �����̏ڍו\��
- **Summary**: �t�@�C���ꗗ�ƕύX��
- **�ʃt�@�C���V�[�g**: �e�t�@�C���̏ڍׂȍ���

## �v���W�F�N�g�\��

```
winmerge-diff-exporter/
������ main.py                    # ���C���G���g���[�|�C���g
������ gui.py                     # GUI����
������ winmergexlsx.py           # WinMerge�A�g��Excel�ϊ�
������ diffdetailsheetcreater.py # �����ڍ׃V�[�g�쐬
������ common.py                  # ���ʊ֐��ƃ��[�e�B���e�B
������ requirements.txt           # Python�ˑ��֌W
������ qt.conf                   # Qt�ݒ�t�@�C��
������ README.md                 # ���̃t�@�C��
������ LICENSE                   # MIT���C�Z���X
```

## �Z�p�d�l

- **GUI �t���[�����[�N**: PyQt6
- **Excel����**: openpyxl + pywin32 (COM�o�R)
- **������r**: WinMerge (�O���v���Z�X)
- **�}���`�X���b�h**: QThread�g�p��UI�u���b�N�����


## ���C�Z���X

MIT License - �ڍׂ�[LICENSE](LICENSE)�t�@�C�����Q�Ƃ��Ă��������B

## �v��

�o�O�񍐂�@�\��Ă�[Issues](https://github.com/your-username/winmerge-diff-exporter/issues)�ł��肢���܂��B

## �ӎ�

���̃\�t�g�E�F�A��[winmerge_xlsx](https://github.com/y-tetsu/winmerge_xlsx.git)�̃R�[�h���܂�ł��܂��B
