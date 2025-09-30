# -*- coding: utf-8 -*-
"""
Common utilities and base classes
"""

import time
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional


class Timer:
    """Performance timer for measuring execution time"""
    
    def __init__(self, label: str = "Processing"):
        self.label = label
        self._sessions = []
        self._memos = []
    
    def start(self, memo: Optional[str] = None) -> None:
        """Start timing a new session"""
        self._sessions.append({'start': time.time(), 'end': None})
        indent = '_' * (len(self._sessions) * 2)
        self._memos.append(memo or "")
        print(f"{indent}{self.label} {memo} started...")

    def stop(self) -> None:
        """Stop timing the current session"""
        if not self._sessions or self._sessions[-1]['end'] is not None:
            print(f"{self.label} has not been started yet.")
            return

        self._sessions[-1]['end'] = time.time()
        elapsed = self._sessions[-1]['end'] - self._sessions[-1]['start']
        indent = '_' * (len(self._sessions) * 2)
        memo = self._memos[-1]
        print(f"{indent}{self.label} {memo} completed. Elapsed time: {elapsed:.2f} seconds")

        # Remove completed session
        self._sessions.pop()
        self._memos.pop()

    def elapsed_all(self) -> float:
        """Calculate total elapsed time across all sessions"""
        return sum(
            (s['end'] or time.time()) - s['start']
            for s in self._sessions
        )


class Logger:
    """Centralized logging handler with console and file output"""
    
    def __init__(
        self,
        name: str = "WinMergeDiffExporter",
        level: int = logging.INFO,
        log_file: Optional[Path] = None,
        max_bytes: int = 1_000_000,  # 1MB
        backup_count: int = 3
    ):
        """
        Initialize logger with console and optional file handlers
        
        Args:
            name: Logger name
            level: Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
            log_file: Optional file path for log output
            max_bytes: Maximum size of log file before rotation
            backup_count: Number of backup log files to keep
        """
        self.logger = logging.getLogger(name)
        self.logger.setLevel(level)
        self.logger.propagate = False

        if not self.logger.handlers:
            formatter = logging.Formatter(
                "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S"
            )

            # Console handler
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)

            # File handler with rotation
            if log_file:
                log_file = Path(log_file)
                log_file.parent.mkdir(parents=True, exist_ok=True)
                file_handler = RotatingFileHandler(
                    log_file, maxBytes=max_bytes, backupCount=backup_count, encoding="utf-8"
                )
                file_handler.setFormatter(formatter)
                self.logger.addHandler(file_handler)

    def debug(self, message: str, exc_info: bool = False) -> None:
        """Log debug message"""
        self.logger.debug(message, exc_info=exc_info)

    def info(self, message: str, exc_info: bool = False) -> None:
        """Log info message"""
        self.logger.info(message, exc_info=exc_info)

    def warning(self, message: str, exc_info: bool = False) -> None:
        """Log warning message"""
        self.logger.warning(message, exc_info=exc_info)

    def error(self, message: str, exc_info: bool = False) -> None:
        """Log error message"""
        self.logger.error(message, exc_info=exc_info)

    def critical(self, message: str, exc_info: bool = False) -> None:
        """Log critical message"""
        self.logger.critical(message, exc_info=exc_info)


# Global logger instance
logger = Logger()
