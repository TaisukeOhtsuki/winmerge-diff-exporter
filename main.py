# -*- coding: UTF-8 -*-
import sys
from PyQt6.QtWidgets import QApplication
from gui import DiffApp

if __name__ == "__main__":
    print("loading now....")
    app = QApplication(sys.argv)
    window = DiffApp()
    window.show()
    sys.exit(app.exec())