# os_compatibility.py
import os
import platform
import logging
import sys
import subprocess
from typing import Dict, Any, Optional

# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

class OSCompatibilityManager:
    """
    Centralized manager for OS-specific behaviors and configurations.
    Focused on Windows 10 vs Windows 11 differences.
    """
    # OS detection results (cached)
    _is_windows_11 = None
    _platform_info = None

    # Timeout multipliers
    TIMEOUT_MULTIPLIERS = {
        'windows_10': 1.0,
        'windows_11': 1.5,  # Windows 11 gets 50% longer timeouts
    }

    # Base timeouts in seconds
    BASE_TIMEOUTS = {
        "browse_dialog": 5,
        "preferences": 3,
        "copy_files": 3,
        "error_window": 5,
        "call_bases": 5,
        "process_completion": 30,
        "read_info": 3,
        "first_browse_delay": 1.2,
        "subsequent_browse_delay": 0.5,
        "click_delay": 0.2,
        "expand_delay": 0.3,
        "polling_interval": 0.5
    }
        
    @classmethod
    def py32_check(cls, script_path=None, logger=None, use_venv=True):
        """
        Check if running in 64-bit Python and restart in 32-bit if needed.
        
        Args:
            script_path: Path to the script being run
            logger: Logger for output
            use_venv: Whether to activate the virtual environment
        """
        # Check if we're already in a relaunch attempt
        if os.environ.get('PYTHON_LAUNCHER_ACTIVE') == '1':
            if logger:
                logger.info("Already in a relaunch attempt, continuing with current interpreter")
            return True
            
        # Check if running in 64-bit Python
        if sys.maxsize > 2 ** 32:
            # Get 32-bit Python path
            py32_path = os.environ.get('PYTHON32_PATH') or getattr(
                __import__('mseqauto.config', fromlist=['MseqConfig']).MseqConfig, 
                'PYTHON32_PATH', 
                None
            )
            
            # Get venv path if needed
            venv_path = None
            venv_activate = None
            if use_venv:
                config = __import__('mseqauto.config', fromlist=['MseqConfig']).MseqConfig
                venv_path = getattr(config, 'VENV32_PATH', 
                                os.path.join(os.path.dirname(py32_path), "venv32")) #type: ignore
                venv_activate = os.path.join(venv_path, "Scripts", "activate.bat")
            
            # Only relaunch if we have a valid 32-bit Python path
            if (script_path and py32_path and os.path.exists(py32_path) 
                    and py32_path != sys.executable):
                
                # Set the environment variable to prevent recursive relaunching
                os.environ['PYTHON_LAUNCHER_ACTIVE'] = '1'
                
                if logger:
                    logger.info(f"Restarting with 32-bit Python: {py32_path}")
                    
                # Use batch file approach for proper path handling
                if use_venv and os.path.exists(venv_path): #type: ignore
                    # Create and run a batch file for venv activation
                    batch_file = os.path.join(os.environ.get('TEMP', os.getcwd()), 
                                            'run_script.bat')
                    with open(batch_file, 'w') as f:
                        f.write(f'@echo off\n')
                        f.write(f'call "{venv_activate}"\n')
                        f.write(f'"{py32_path}" "{script_path}"')
                        # Add arguments if any
                        if len(sys.argv) > 1:
                            f.write(f' {" ".join(sys.argv[1:])}')
                        f.write('\n')
                    
                    # Run the batch file
                    os.system(batch_file)
                    
                    # Clean up
                    try:
                        os.remove(batch_file)
                    except:
                        pass
                else:
                    # Direct execution without venv
                    subprocess.run([py32_path, script_path] + sys.argv[1:])
                    
                sys.exit(0)
                
            else:
                if logger:
                    logger.info("32-bit Python not found or is current interpreter")
                    
        return True

    @classmethod
    def is_windows_11(cls) -> bool:
        """Detect if the system is running Windows 11."""
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
        """Get timeout value adjusted for the current OS."""
        # Get base timeout from constants or override
        base_timeout = base_override if base_override is not None else cls.BASE_TIMEOUTS.get(operation_type, 5.0)

        # Get multiplier based on OS
        multiplier = cls.TIMEOUT_MULTIPLIERS.get(cls.get_os_key(), 1.0)
        
        # Environment variable can override
        env_multiplier = os.environ.get('MSEQ_TIMEOUT_MULTIPLIER')
        if env_multiplier:
            try:
                multiplier = float(env_multiplier)
            except ValueError:
                pass

        return base_timeout * multiplier

    @classmethod
    def log_environment_info(cls, logger: Optional[logging.Logger] = None):
        """Log information about the operating environment."""
        if logger is None:
            logger = logging.getLogger(__name__)

        info = cls.get_platform_info()
        logger.info("OS Environment Information:")
        logger.info(f"  System: {info['system']} {info['release']} ({info['version']})")
        logger.info(f"  Windows 11: {'Yes' if info['is_windows_11'] else 'No'}")
        logger.info(f"  Architecture: {info['architecture'][0]} ({info['machine']})")
        logger.info(f"  OS Key: {cls.get_os_key()}")
        logger.info(f"  Timeout Multiplier: {cls.TIMEOUT_MULTIPLIERS.get(cls.get_os_key(), 1.0)}")
        logger.info(f"  Python: {platform.python_version()} ({platform.architecture()[0]})")
        logger.info(f"  Interpreter: {sys.executable}")