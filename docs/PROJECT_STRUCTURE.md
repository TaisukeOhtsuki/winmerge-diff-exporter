# Project Structure

## Directory Layout

```
winmerge-diff-exporter/
������ main.py                 # Application entry point
������ requirements.txt        # Python dependencies
������ qt.conf                 # Qt configuration
������ LICENSE                 # License file
������ README.md              # Project documentation
��
������ src/                   # Source code
��   ������ __init__.py
��   ������ core/              # Core business logic
��   ��   ������ __init__.py
��   ��   ������ common.py      # Logger and Timer utilities
��   ��   ������ config.py      # Configuration management
��   ��   ������ exceptions.py  # Custom exceptions
��   ��   ������ utils.py       # File operations and Excel formatting
��   ��   ������ winmergexlsx.py           # WinMerge integration
��   ��   ������ diffdetailsheetcreater.py # Diff detail sheet creator
��   ��
��   ������ converters/        # File converters
��   ��   ������ __init__.py
��   ��   ������ html_to_excel.py  # HTML to Excel converter
��   ��
��   ������ ui/                # User interface
��       ������ __init__.py
��       ������ gui.py         # PyQt6 GUI application
��
������ docs/                  # Documentation
��   ������ EXCEL_COM_REMOVAL.md
��   ������ FILE_LOCK_COMPLETE_FIX.md
��   ������ FREEZE_FIX.md
��   ������ FREEZE_FIX_SUMMARY.md
��   ������ IMPROVEMENT_SUMMARY.md
��   ������ PERMISSION_ERROR_FIX.md
��   ������ REFACTORING_SUMMARY.md
��   ������ RELEASE_NOTES_v2.0.md
��
������ tests/                 # Unit tests
��   ������ __init__.py
��
������ output/                # Output files
��   ������ output.html
��   ������ output.xlsx
��   ������ output.files/
��
������ venv/                  # Python virtual environment
��
������ folder1/, folder2/     # Test folders for comparison

## Module Dependencies

```
main.py
  ������ src.ui.gui
      ������ src.core.winmergexlsx
          ������ src.core.config
          ������ src.core.exceptions
          ������ src.core.utils
          ������ src.core.common
          ������ src.core.diffdetailsheetcreater
          ������ src.converters.html_to_excel
```

## Key Features by Module

### Core Modules (`src/core/`)
- **common.py**: Logging and timing utilities
- **config.py**: Centralized configuration with validation
- **exceptions.py**: Custom exception classes
- **utils.py**: File operations, Excel formatting, path management
- **winmergexlsx.py**: Main logic for WinMerge integration and Excel conversion
- **diffdetailsheetcreater.py**: Creates detailed comparison sheets

### Converters (`src/converters/`)
- **html_to_excel.py**: Converts HTML tables to Excel (no COM dependencies)

### UI (`src/ui/`)
- **gui.py**: PyQt6-based drag-and-drop interface with threading

## Installation

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
venv\Scripts\activate  # Windows

# Install dependencies
pip install -r requirements.txt
```

## Usage

```bash
# Run the application
python main.py
```

## Development

### Adding New Features
1. Core business logic goes in `src/core/`
2. UI components go in `src/ui/`
3. File converters go in `src/converters/`
4. Tests go in `tests/`
5. Documentation goes in `docs/`

### Code Organization Principles
- **Separation of Concerns**: UI, business logic, and utilities are separated
- **Modular Design**: Each module has a single responsibility
- **Relative Imports**: Use relative imports within packages (`.module`)
- **Absolute Imports**: Use absolute imports from outside (`src.package.module`)
