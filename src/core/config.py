# -*- coding: UTF-8 -*-
"""
Application configuration settings
"""
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, List, Any


@dataclass
class WinMergeConfig:
    """WinMerge configuration"""
    executable_path: str = r'C:\Program Files\WinMerge\WinMergeU.exe'
    options: List[str] = field(default_factory=list)
    
    def __post_init__(self):
        if not self.options:
            self.options = [
                '/minimize',                                    # Run minimized
                '/noninteractive',                              # Exit after generating report
                '/cfg', 'Settings/DirViewExpandSubdirs=1',      # Auto-expand subdirectories
                '/cfg', 'ReportFiles/ReportType=2',             # Simple HTML format
                '/cfg', 'ReportFiles/IncludeFileCmpReport=1',   # Include file comparison report
                '/cfg', 'Settings/ViewLineNumbers=1',           # Show line numbers
                '/cfg', 'Settings/IgnoreEol=1',                 # Ignore line ending differences
                '/r',                                           # Compare all files in subdirectories
                '/u',                                           # Don't add to recent list
                '/or',                                          # Output report
            ]


@dataclass
class ExcelConfig:
    """Excel formatting configuration"""
    summary_ws_num: int = 1
    summary_start_row: int = 6
    summary_name_col_index: int = 1
    summary_folder_col_index: int = 2
    diff_start_row: int = 2
    diff_zoom_ratio: int = 85
    home_position: str = 'A1'
    
    # Excel constants
    xl_up: int = -4162
    xl_open_xml_workbook: int = 51
    xl_center: int = -4108
    xl_continuous: int = 1


@dataclass 
class UIConfig:
    """UI configuration"""
    window_title: str = "WinMerge Diff to Excel"
    window_geometry: tuple = (100, 100, 800, 500)
    default_output_file: str = "output.xlsx"
    progress_animation_interval: int = 50
    progress_animation_step: int = 2


@dataclass
class DiffConfig:
    """Diff processing configuration"""
    context_lines: int = 4
    yellow_color: str = 'FFC0C0C0'
    sheet_start_index: int = 2


class Config:
    """Global configuration manager"""
    
    def __init__(self):
        self.winmerge = WinMergeConfig()
        self.excel = ExcelConfig()
        self.ui = UIConfig()
        self.diff = DiffConfig()
        
        # Diff formats configuration
        self.diff_formats = {
            'no': [
                {'col': 'A', 'width': 5},
                {'col': 'C', 'width': 5},
            ],
            'code': [
                {'col': 'B', 'width': 100, 'font': 'MS Gothic', 'header': 'Before'},
                {'col': 'D', 'width': 100, 'font': 'MS Gothic', 'header': 'After'},
            ],
            'extra': [
                {'col': 'E', 'width': 60, 'comment': 'Comments'},
            ],
        }
    
    def validate(self) -> bool:
        """Validate configuration settings"""
        winmerge_path = Path(self.winmerge.executable_path)
        if not winmerge_path.exists():
            return False
        return True
    
    def get_winmerge_command(self, base: str, latest: str, output: str) -> List[str]:
        """Build WinMerge command with all options"""
        return [
            self.winmerge.executable_path,
            str(base),
            str(latest),
            *self.winmerge.options,
            str(output),
        ]


# Global configuration instance
config = Config()
