# Project Structure

## Directory Layout

```
winmerge-diff-exporter/
„¥„Ÿ„Ÿ main.py                 # Application entry point
„¥„Ÿ„Ÿ requirements.txt        # Python dependencies
„¥„Ÿ„Ÿ qt.conf                 # Qt configuration
„¥„Ÿ„Ÿ LICENSE                 # License file
„¥„Ÿ„Ÿ README.md              # Project documentation
„ 
„¥„Ÿ„Ÿ src/                   # Source code
„    „¥„Ÿ„Ÿ __init__.py
„    „¥„Ÿ„Ÿ core/              # Core business logic
„    „    „¥„Ÿ„Ÿ __init__.py
„    „    „¥„Ÿ„Ÿ common.py      # Logger and Timer utilities
„    „    „¥„Ÿ„Ÿ config.py      # Configuration management
„    „    „¥„Ÿ„Ÿ exceptions.py  # Custom exceptions
„    „    „¥„Ÿ„Ÿ utils.py       # File operations and Excel formatting
„    „    „¥„Ÿ„Ÿ winmergexlsx.py           # WinMerge integration
„    „    „¤„Ÿ„Ÿ diffdetailsheetcreater.py # Diff detail sheet creator
„    „ 
„    „¥„Ÿ„Ÿ converters/        # File converters
„    „    „¥„Ÿ„Ÿ __init__.py
„    „    „¤„Ÿ„Ÿ html_to_excel.py  # HTML to Excel converter
„    „ 
„    „¤„Ÿ„Ÿ ui/                # User interface
„        „¥„Ÿ„Ÿ __init__.py
„        „¤„Ÿ„Ÿ gui.py         # PyQt6 GUI application
„ 
„¥„Ÿ„Ÿ docs/                  # Documentation
„    „¥„Ÿ„Ÿ EXCEL_COM_REMOVAL.md
„    „¥„Ÿ„Ÿ FILE_LOCK_COMPLETE_FIX.md
„    „¥„Ÿ„Ÿ FREEZE_FIX.md
„    „¥„Ÿ„Ÿ FREEZE_FIX_SUMMARY.md
„    „¥„Ÿ„Ÿ IMPROVEMENT_SUMMARY.md
„    „¥„Ÿ„Ÿ PERMISSION_ERROR_FIX.md
„    „¥„Ÿ„Ÿ REFACTORING_SUMMARY.md
„    „¤„Ÿ„Ÿ RELEASE_NOTES_v2.0.md
„ 
„¥„Ÿ„Ÿ tests/                 # Unit tests
„    „¤„Ÿ„Ÿ __init__.py
„ 
„¥„Ÿ„Ÿ output/                # Output files
„    „¥„Ÿ„Ÿ output.html
„    „¥„Ÿ„Ÿ output.xlsx
„    „¤„Ÿ„Ÿ output.files/
„ 
„¥„Ÿ„Ÿ venv/                  # Python virtual environment
„ 
„¤„Ÿ„Ÿ folder1/, folder2/     # Test folders for comparison

## Module Dependencies

```
main.py
  „¤„Ÿ„Ÿ src.ui.gui
      „¤„Ÿ„Ÿ src.core.winmergexlsx
          „¥„Ÿ„Ÿ src.core.config
          „¥„Ÿ„Ÿ src.core.exceptions
          „¥„Ÿ„Ÿ src.core.utils
          „¥„Ÿ„Ÿ src.core.common
          „¥„Ÿ„Ÿ src.core.diffdetailsheetcreater
          „¤„Ÿ„Ÿ src.converters.html_to_excel
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
