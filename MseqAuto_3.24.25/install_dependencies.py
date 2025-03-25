import sys
import os
import subprocess
import platform

def install_dependencies():
    """Install all required Python packages"""
    print("Installing dependencies for MseqAuto...")
    
    dependencies = [
        "pywinauto>=0.6.8",
        "numpy>=1.20.0",
        "openpyxl>=3.0.7", 
        "pylightxl>=1.60",
        "pywin32>=300",
        "psutil>=5.9.0",  # Process and system monitoring
        "pyinstaller>=5.7.0"  # Useful for creating standalone executables if needed
    ]
    
    # Create requirements.txt
    with open("requirements.txt", "w") as f:
        f.write("\n".join(dependencies))
    
    # Install packages
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], check=True)
        print("All dependencies installed successfully.")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error installing dependencies: {e}")
        return False

def create_logs_directory():
    """Create logs directory if it doesn't exist"""
    print("Creating logs directory...")
    if not os.path.exists("logs"):
        os.makedirs("logs")
        print("Logs directory created.")
    else:
        print("Logs directory already exists.")

def main():
    print("=== MseqAuto Setup ===")
    print(f"Python version: {platform.python_version()} ({64 if sys.maxsize > 2**32 else 32}-bit)")
    
    # Install dependencies
    deps_ok = install_dependencies()
    if not deps_ok:
        print("ERROR: Failed to install dependencies.")
        print("Setup aborted.")
        return
    
    # Create logs directory
    create_logs_directory()
    
    print("\nSetup completed successfully!")
    print("You can now run the MseqAuto scripts.")

if __name__ == "__main__":
    main()