"""
Core functionality for the MseqAuto package.

This module contains the core classes and functions that handle
file system operations, folder processing, and UI automation.
"""

from .file_system_dao import FileSystemDAO
from .folder_processor import FolderProcessor
from .ui_automation import MseqAutomation

__all__ = ['FileSystemDAO', 'FolderProcessor', 'MseqAutomation']