# -*- coding: UTF-8 -*-
import sys
import os
import ctypes

# Suppress Qt DPI warnings only
os.environ["QT_LOGGING_RULES"] = "qt.qpa.window.debug=false;qt.qpa.window.warning=false"
os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"  # Enable proper scaling
os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

# Set system DPI aware to maintain readability but avoid context conflicts
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

from PyQt6.QtWidgets import QApplication
from gui import DiffApp


if __name__ == "__main__":
    print("loading now....")
    app = QApplication(sys.argv)
    window = DiffApp()
    window.show()
    sys.exit(app.exec())