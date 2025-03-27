"""
Core functionality for the MseqAuto package.

This module contains the core classes and functions that handle
file system operations, folder processing, and UI automation.
"""

# First, import the OSCompatibilityManager to make it available
from .os_compatibility import OSCompatibilityManager

# Then import other classes that might depend on it
from .file_system_dao import FileSystemDAO
from .folder_processor import FolderProcessor
from .ui_automation import MseqAutomation

__all__ = ['FileSystemDAO', 'FolderProcessor', 'MseqAutomation', 'OSCompatibilityManager']
