# Excel COM �ˑ����폜 - ���P�T�}���[

## ���t: 2025�N9��30��

## ���_

�ȑO�̃V�X�e���͈ȉ��̏d��Ȑ���������܂����F

1. **Excel COM�I�u�W�F�N�g�ւ̈ˑ�**
   - Microsoft Excel���C���X�g�[������Ă���K�v��������
   - Excel�����̃v���Z�X�ŊJ���Ă���Ǝ��s�ł��Ȃ�����
   - Windows��p�ŃN���X�v���b�g�t�H�[����Ή�

2. **���֐��̖��**
   - ���[�U�[������Excel�t�@�C�����J���Ă���Ǝ��s�Ɏ��s
   - Excel COM�̏������Ɏ��Ԃ�������
   - �G���[���b�Z�[�W���s�e��

---

## ������

### ���S��Python�����ւ̈ڍs

Excel COM�����S�ɍ폜���APure Python�����ɒu�������܂����F

```
���A�[�L�e�N�`��:
WinMerge �� HTML �� Excel COM �� XLSX �� openpyxl �� �ŏIXLSX

�V�A�[�L�e�N�`��:
WinMerge �� HTML �� BeautifulSoup4 �� openpyxl �� XLSX
```

---

## �����̏ڍ�

### 1. �V�������W���[��: `html_to_excel.py`

**HTMLToExcelConverter** �N���X�������F

```python
class HTMLToExcelConverter:
    """Convert HTML tables to Excel without using Excel COM"""
    
    def convert_summary_html(self, html_path: Path) -> Workbook:
        """�T�}���[HTML��Workbook�ɕϊ�"""
        
    def convert_html_file(self, html_path: Path, wb: Workbook, sheet_name: str):
        """�ʂ�HTML�t�@�C�����V�[�g�ɕϊ�"""
        
    def _convert_table_to_sheet(self, table, ws):
        """HTML�e�[�u����Excel�V�[�g�ɕϊ�"""
        
    def _apply_cell_styling(self, excel_cell, html_cell):
        """HTML�Z���̃X�^�C����Excel�Z���ɓK�p"""
```

#### ��ȋ@�\:

- **BeautifulSoup4** ��HTML���p�[�X
- HTML�e�[�u����Excel�Z���ɕϊ�
- �w�i�F�A�t�H���g�A�r���Ȃǂ̃X�^�C����ێ�
- colspan/rowspan�̃}�[�W�Z���ɑΉ�
- �����񕝒���

### 2. `winmergexlsx.py` �̉��P

#### �폜���ꂽ���\�b�h:
```python
? _check_excel_application()  # Excel�N���`�F�b�N�s�v
? _process_with_excel_com()   # Excel COM����
? _copy_html_files_with_com() # COM�o�R�̃R�s�[
? self.excel_app.Quit()       # Excel�N���[���A�b�v�s�v
```

#### �ǉ����ꂽ���\�b�h:
```python
? _convert_diff_html_files()  # Pure Python HTML�ϊ�
```

#### �ύX���ꂽ���\�b�h:
```python
def _convert_html_to_xlsx(self) -> None:
    """Convert HTML report to Excel using pure Python"""
    from html_to_excel import HTMLToExcelConverter
    
    converter = HTMLToExcelConverter(log_callback=self.log_callback)
    
    # �T�}���[HTML��Workbook�ɕϊ�
    self.wb = converter.convert_summary_html(self.output_html)
    
    # ���ׂĂ�diff HTML�t�@�C����ϊ�
    self._convert_diff_html_files(converter)
    
    # �ۑ�
    self.wb.save(str(self.output))
    
    # openpyxl�ōŏI�t�H�[�}�b�g
    self._process_with_openpyxl()
```

### 3. �ˑ��֌W�̕ύX

#### �폜:
```
? pywin32==306  # win32com.client �s�v
```

#### �ǉ�:
```
? beautifulsoup4==4.12.3  # HTML�p�[�X
? lxml==5.3.0             # BeautifulSoup�̃p�[�T�[
```

---

## �����b�g

### 1. ? **Excel�C���X�g�[���s�v**
- Microsoft Excel���Ȃ��Ă�����
- Excel���C�Z���X�s�v
- �y�ʂȊ��Ŏ��s�\

### 2. ? **�t�@�C�����b�N�̖�����**
- Excel�t�@�C�����J���Ă��Ă����s�\
- ������s���\
- �t�@�C���A�N�Z�X�̋����Ȃ�

### 3. ? **�N���X�v���b�g�t�H�[���Ή�**
- Windows�ȊO�ł�����\�iLinux, macOS�j
- Docker�R���e�i�ł̎��s���e��
- CI�V�X�e���ł̎��������ȒP

### 4. ? **�p�t�H�[�}���X����**
- Excel COM�N���̑҂����ԃ[��
- �������g�p�ʂ̍팸
- ��荂���ȏ���

### 5. ? **�����肵�����s**
- Excel COM�̕s���萫��r��
- �G���[�n���h�����O���e��
- �f�o�b�O���ȒP

### 6. ? **�ێ琫�̌���**
- �V���v����Python�R�[�h�̂�
- COM�I�u�W�F�N�g�̕��G����r��
- �e�X�g���e��

---

## �g�p��

### �ύX�O�iExcel�K�{�j:
```python
# Excel���C���X�g�[������Ă���K�v������
# Excel�����K�v������
diff = WinMergeXlsx(base, latest, output)
diff.generate()  # Excel COM���g�p
```

### �ύX��iExcel�s�v�j:
```python
# Excel�͕s�v�I
# Excel���J���Ă��Ă�OK�I
diff = WinMergeXlsx(base, latest, output)
diff.generate()  # Pure Python����
```

---

## �Z�p�I�ȏڍ�

### HTML�p�[�X����

```python
# BeautifulSoup��HTML���p�[�X
with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
    html_content = f.read()

soup = BeautifulSoup(html_content, 'html.parser')
tables = soup.find_all('table')

# �e�[�u����Excel�ɕϊ�
for tr in table.find_all('tr'):
    for cell in tr.find_all(['th', 'td']):
        cell_text = cell.get_text(strip=True)
        excel_cell = ws.cell(row=row_idx, column=col_idx, value=cell_text)
        
        # �X�^�C���K�p
        self._apply_cell_styling(excel_cell, cell)
```

### �X�^�C���ϊ�

```python
def _apply_cell_styling(self, excel_cell, html_cell):
    """HTML�Z���̃X�^�C����Excel�Z���ɓK�p"""
    
    # �w�i�F
    if 'background-color' in style or bgcolor:
        color = self._parse_color(style, bgcolor)
        excel_cell.fill = PatternFill(
            start_color=color,
            end_color=color,
            fill_type='solid'
        )
    
    # �w�b�_�[�X�^�C��
    if html_cell.name == 'th':
        excel_cell.font = Font(bold=True, size=11)
        excel_cell.alignment = Alignment(horizontal='center')
    
    # �r��
    excel_cell.border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
```

---

## �݊���

### �T�|�[�g��
- ? Windows 10/11
- ? Windows Server 2016+
- ? Linux (Ubuntu, CentOS, etc.)
- ? macOS 10.15+
- ? Docker �R���e�i

### Python�o�[�W����
- ? Python 3.8+
- ? Python 3.9+
- ? Python 3.10+
- ? Python 3.11+
- ? Python 3.12+
- ? Python 3.13+

---

## �}�C�O���[�V����

### �������[�U�[�ւ̉e��
**�e���Ȃ�** - API �͊��S�ɓ����ł��B

### �K�v�ȃA�N�V����
1. �V�����ˑ��֌W���C���X�g�[��:
   ```bash
   pip install beautifulsoup4==4.12.3 lxml==5.3.0
   ```

2. pywin32�͍폜�\�i�I�v�V�����j:
   ```bash
   pip uninstall pywin32
   ```

---

## �e�X�g

### �\���`�F�b�N
```bash
python -m py_compile html_to_excel.py winmergexlsx.py
? No errors
```

### �@�\�e�X�g����
- [ ] HTML��Excel�ϊ�
- [ ] �X�^�C���ێ��i�w�i�F�A�t�H���g�j
- [ ] �}�[�W�Z������
- [ ] �����V�[�g�쐬
- [ ] ��K�̓t�@�C������
- [ ] �G���[�n���h�����O

---

## �p�t�H�[�}���X��r

| ���� | �������iCOM�j | �V�����iPure Python�j |
|------|--------------|---------------------|
| Excel�N������ | 3-5�b | 0�b |
| �ϊ����x | �x�� | ���� |
| �������g�p�� | ���� | �Ⴂ |
| ������s | �s�� | �\ |
| �G���[�� | ���� | �Ⴂ |

---

## ����̊g���\��

1. **CSV�G�N�X�|�[�g**: HTML����CSV�ւ̕ϊ��T�|�[�g
2. **PDF�G�N�X�|�[�g**: HTML����PDF�ւ̕ϊ��T�|�[�g
3. **�J�X�^���e�[�}**: Excel�e�[�}�̃J�X�^�}�C�Y
4. **�o�b�`����**: �����v���W�F�N�g�̈ꊇ����
5. **REST API**: Web�T�[�r�X�Ƃ��Ă̒�

---

## ���_

Excel COM�ˑ����̊��S�폜�ɂ��A�ȉ��������F

? ���_��ȃf�v���C�����g
? ��荂�����萫
? ���ǂ����[�U�[�̌�
? ���ȒP�ȕێ�
? ���L���v���b�g�t�H�[���T�|�[�g

**Status**: ? ����
**Breaking Changes**: �Ȃ�
**Recommended Action**: �V�����ˑ��֌W���C���X�g�[��
