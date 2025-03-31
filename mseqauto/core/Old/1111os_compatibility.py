# os_compatibility.py
import os
import platform
import logging
import sys
import subprocess
from typing import Dict, Any, Optional
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
print(sys.path)
from mseqauto.config import MseqConfig

class OSCompatibilityManager:
    """
    Centralized manager for OS-specific behaviors and configurations.
    Primarily focused on Windows 10 vs Windows 11 differences.
    """
    # OS detection results (cached)
    _is_windows_11 = None
    _platform_info = None

    # Timeout multipliers relative to base timeouts
    TIMEOUT_MULTIPLIERS = {
        'windows_10': 1.0,
        'windows_11': 1.5,  # Windows 11 gets 50% longer timeouts by default
    }

    # Base timeouts in seconds - centralized here for easy modification
    BASE_TIMEOUTS = {
        "browse_dialog": 5,
        "preferences": 5,
        "copy_files": 5,
        "error_window": 10,
        "call_bases": 10,
        "process_completion": 45,
        "read_info": 5,
        "navigation": 5,
        "window_appear": 2,
        "click_response": 1,
        "polling_interval": 0.5
    }

    # Dialog behaviors - difference in Windows 11 dialog handling
    DIALOG_STRATEGIES = {
        'windows_10': {
            'browse_dialog': {
                'prefer_tree_navigation': True,
                'click_location': 'center',
                'use_fallback_buttons': False,
            },
        },
        'windows_11': {
            'browse_dialog': {
                'prefer_tree_navigation': True,
                'click_location': 'top_center',  # Windows 11 has different control layouts
                'use_fallback_buttons': True,
                'extra_sleep_after_click': 0.5,  # Extra delay for Win11 animation
            },
        }
    }
        
    @classmethod
    def py32_check(cls, script_path=None, logger=None):
        """
        Check if running in 64-bit Python and restart in 32-bit if needed.

        Args:
            script_path: Path to the script to restart (must be provided)
            logger: Optional logger for recording the outcome

        Returns:
            bool: True if check passed (no restart needed), False otherwise
        """
        import sys
        import os
        import subprocess

        # No script path provided, can't restart properly
        if not script_path:
            if logger:
                logger.warning("No script path provided for 32-bit Python check")
            return False

        # Check if running in 64-bit Python
        if sys.maxsize > 2 ** 32:
            # Path to 32-bit Python
            from mseqauto.config import MseqConfig
            py32_path = MseqConfig.PYTHON32_PATH

            # Only relaunch if we have a different 32-bit Python
            if os.path.exists(py32_path) and py32_path != sys.executable:
                if logger:
                    logger.info(f"Restarting with 32-bit Python: {py32_path}")

                # Re-run the script with 32-bit Python and exit current process
                subprocess.run([py32_path, script_path])
                sys.exit(0)
            else:
                if logger:
                    logger.info("32-bit Python not specified or same as current interpreter")
                    logger.info("Continuing with current Python interpreter")

        # If we get here, no restart was needed or possible
        return True

    @classmethod
    def is_windows_11(cls) -> bool:
        """
        Detect if the system is running Windows 11.
        Caches result for efficiency.
        """
        if cls._is_windows_11 is None:
            if platform.system() != 'Windows':
                cls._is_windows_11 = False
            else:
                # Windows 11 is Windows 10 with build >= 22000
                win_version = int(platform.version().split('.')[0])
                win_build = int(platform.version().split('.')[2]) if len(platform.version().split('.')) > 2 else 0
                cls._is_windows_11 = (win_version >= 10 and win_build >= 22000)

        return cls._is_windows_11

    @classmethod
    def get_platform_info(cls) -> Dict[str, Any]:
        """Get detailed platform information."""
        if cls._platform_info is None:
            cls._platform_info = {
                'system': platform.system(),
                'release': platform.release(),
                'version': platform.version(),
                'is_windows_11': cls.is_windows_11(),
                'architecture': platform.architecture(),
                'machine': platform.machine()
            }

        return cls._platform_info

    @classmethod
    def get_os_key(cls) -> str:
        """Get the key to use for OS-specific configurations."""
        return 'windows_11' if cls.is_windows_11() else 'windows_10'

    @classmethod
    def get_timeout(cls, operation_type: str, base_override: Optional[float] = None) -> float:
        """
        Get the appropriate timeout value for an operation based on OS.

        Args:
            operation_type: The type of operation (e.g., "browse_dialog")
            base_override: Optional override for the base timeout value

        Returns:
            The adjusted timeout value in seconds
        """
        # Get base timeout from constants or override
        base_timeout = base_override if base_override is not None else cls.BASE_TIMEOUTS.get(operation_type, 5.0)

        # Get multiplier based on OS
        os_key = cls.get_os_key()
        multiplier = cls.TIMEOUT_MULTIPLIERS.get(os_key, 1.0)

        # Environment variable can override multiplier
        env_multiplier = os.environ.get('MSEQ_TIMEOUT_MULTIPLIER')
        if env_multiplier:
            try:
                multiplier = float(env_multiplier)
            except ValueError:
                pass

        return base_timeout * multiplier

    @classmethod
    def get_dialog_strategy(cls, dialog_type: str) -> Dict[str, Any]:
        """
        Get the appropriate dialog handling strategy based on OS.

        Args:
            dialog_type: The type of dialog (e.g., "browse_dialog")

        Returns:
            Dictionary of strategy parameters
        """
        os_key = cls.get_os_key()
        os_strategies = cls.DIALOG_STRATEGIES.get(os_key, {})

        # Return strategy for this dialog type, or empty dict if not found
        return os_strategies.get(dialog_type, {})

    @classmethod
    def log_environment_info(cls, logger: Optional[logging.Logger] = None):
        """
        Log detailed information about the operating environment.
        Useful for diagnostics.
        """
        if logger is None:
            logger = logging.getLogger(__name__)

        info = cls.get_platform_info()
        logger.info("OS Environment Information:")
        logger.info(f"  System: {info['system']} {info['release']} ({info['version']})")
        logger.info(f"  Windows 11: {'Yes' if info['is_windows_11'] else 'No'}")
        logger.info(f"  Architecture: {info['architecture'][0]} ({info['machine']})")
        logger.info(f"  OS Key: {cls.get_os_key()}")
        logger.info(f"  Timeout Multiplier: {cls.TIMEOUT_MULTIPLIERS.get(cls.get_os_key(), 1.0)}")

        # Log current Python interpreter
        logger.info(f"  Python: {platform.python_version()} ({platform.architecture()[0]})")
        logger.info(f"  Interpreter: {sys.executable}")


class DialogHandler:
    """
    Handles dialog interactions with OS-specific strategies.
    """

    def __init__(self, app, logger=None):
        self.app = app
        self.logger = logger or logging.getLogger(__name__)
        self.os_manager = OSCompatibilityManager

    def wait_for_dialog(self, dialog_type, timeout=None):
        """
        Wait for a dialog to appear with OS-appropriate timeouts.

        Args:
            dialog_type: Type of dialog to wait for
            timeout: Optional explicit timeout override

        Returns:
            True if dialog appeared, False otherwise
        """
        if timeout is None:
            timeout = self.os_manager.get_timeout(dialog_type)

        try:
            from pywinauto import timings

            # Define the condition based on dialog type
            if dialog_type == "browse_dialog":
                condition = lambda: (self.app.window(title='Browse For Folder').exists() or
                                     self.app.window(title_re='Browse.*Folder').exists())
            elif dialog_type == "preferences":
                condition = lambda: (self.app.window(title='Mseq Preferences').exists() or
                                     self.app.window(title='mSeq Preferences').exists())
            elif dialog_type == "copy_files":
                condition = lambda: self.app.window(title_re='Copy.*sequence files').exists()
            elif dialog_type == "error_window":
                condition = lambda: (self.app.window(title='File error').exists() or
                                     self.app.window(title_re='.*[Ee]rror.*').exists())
            elif dialog_type == "call_bases":
                condition = lambda: self.app.window(title_re='Call bases.*').exists()
            elif dialog_type == "read_info":
                condition = lambda: self.app.window(title_re='Read information for.*').exists()
            else:
                self.logger.warning(f"Unknown dialog type: {dialog_type}")
                return False

            # Wait for the condition with the calculated timeout
            self.logger.info(f"Waiting for {dialog_type} dialog (timeout: {timeout}s)")
            timings.wait_until(timeout=timeout, retry_interval=0.1, func=condition, value=True)
            self.logger.info(f"Dialog {dialog_type} appeared")
            return True
        except Exception as e:
            self.logger.warning(f"Timeout waiting for {dialog_type} dialog: {e}")
            return False