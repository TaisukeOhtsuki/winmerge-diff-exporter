# -*- coding: UTF-8 -*-
"""
Custom exceptions for the application
"""
from typing import Optional


class WinMergeDiffExporterError(Exception):
    """Base exception for WinMerge Diff Exporter"""
    
    def __init__(self, message: str, details: Optional[str] = None):
        super().__init__(message)
        self.message = message
        self.details = details
    
    def __str__(self) -> str:
        if self.details:
            return f"{self.message}\nDetails: {self.details}"
        return self.message


class WinMergeNotFoundError(WinMergeDiffExporterError):
    """Raised when WinMerge executable is not found"""
    pass


class ExcelProcessingError(WinMergeDiffExporterError):
    """Raised when Excel processing fails"""
    pass


class FileProcessingError(WinMergeDiffExporterError):
    """Raised when file processing fails"""
    pass


class ValidationError(WinMergeDiffExporterError):
    """Raised when input validation fails"""
    pass


class ConfigurationError(WinMergeDiffExporterError):
    """Raised when configuration is invalid"""
    pass
