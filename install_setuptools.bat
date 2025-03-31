@echo off
echo ================================================================================
echo Installing required modules for mseqauto...
echo ================================================================================
python -m pip install --upgrade pip
python -m pip install --upgrade setuptools
python -m pip install pywinauto>=0.6.8 numpy>=1.20.0 openpyxl>=3.0.7 pylightxl>=1.60 pywin32>=300 psutil>=5.9.0

echo.
echo Installation complete! Press any key to exit.
pause > nul