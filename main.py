# -*- coding: UTF-8 -*-
"""
WinMerge Diff to Excel Exporter - Main Entry Point
"""
import sys
import os
import ctypes
from typing import NoReturn

from PyQt6.QtWidgets import QApplication

from src.ui.gui import DiffApp
from src.core.common import logger


class DPIConfigurator:
    """Windows DPI awareness configuration handler"""
    
    @staticmethod
    def configure_environment() -> None:
        """Configure environment variables for Qt DPI handling"""
        os.environ["QT_LOGGING_RULES"] = "qt.qpa.window.debug=false;qt.qpa.window.warning=false"
        os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
        os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    
    @staticmethod
    def set_dpi_awareness() -> None:
        """Set Windows DPI awareness to prevent scaling issues"""
        try:
            # Try modern DPI awareness API (Windows 10 1703+)
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
            logger.debug("DPI awareness set via SetProcessDpiAwareness")
        except (AttributeError, OSError):
            try:
                # Fallback to older API (Windows Vista+)
                ctypes.windll.user32.SetProcessDPIAware()
                logger.debug("DPI awareness set via SetProcessDPIAware")
            except (AttributeError, OSError) as e:
                logger.warning(f"Failed to set DPI awareness: {e}")


class ApplicationLauncher:
    """Application launcher with proper initialization"""
    
    def __init__(self):
        self.app: QApplication = None
        self.window: DiffApp = None
    
    def initialize(self) -> None:
        """Initialize application components"""
        logger.info("Initializing application...")
        
        # Configure DPI settings
        DPIConfigurator.configure_environment()
        DPIConfigurator.set_dpi_awareness()
        
        # Create Qt application
        self.app = QApplication(sys.argv)
        self.app.setApplicationName("WinMerge Diff Exporter")
        self.app.setOrganizationName("WinMergeDiffExporter")
        
        # Create main window
        self.window = DiffApp()
        
        logger.info("Application initialized successfully")
    
    def run(self) -> int:
        """Run the application event loop"""
        if not self.app or not self.window:
            raise RuntimeError("Application not initialized. Call initialize() first.")
        
        self.window.show()
        logger.info("Application started")
        
        return self.app.exec()


def main() -> NoReturn:
    """Main entry point"""
    try:
        launcher = ApplicationLauncher()
        launcher.initialize()
        exit_code = launcher.run()
        
        logger.info(f"Application exited with code {exit_code}")
        sys.exit(exit_code)
        
    except Exception as e:
        logger.error(f"Fatal error during application startup: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()