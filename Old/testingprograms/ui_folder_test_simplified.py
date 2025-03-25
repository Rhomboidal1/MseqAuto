from pywinauto import Application
from time import sleep
import io
import sys

# Function to capture print output to a string
def capture_print_output(func, *args, **kwargs):
    old_stdout = sys.stdout
    new_stdout = io.StringIO()
    sys.stdout = new_stdout
    
    try:
        func(*args, **kwargs)
        output = new_stdout.getvalue()
    finally:
        sys.stdout = old_stdout
    
    return output

# Open mSeq
print("Starting mSeq...")
app = Application(backend='win32').start('cmd /c "cd /d C:\\DNA\\Mseq4\\bin && C:\\DNA\\Mseq4\\bin\\j.exe -jprofile mseq.ijl"', wait_for_idle=False).connect(title='mSeq', timeout=10)

# Allow time for window to fully load
print("Waiting for mSeq to initialize...")
sleep(3)

# Get the main window
print("Getting main window...")
try:
    main_window = app.window(title='mSeq')
except:
    try:
        main_window = app.window(title_re='mSeq*')
    except:
        main_window = app.window(title_re='Mseq*')

print(f"Found window: {main_window.window_text()}")

# Capture control identifiers
print("Capturing control identifiers...")
output = capture_print_output(main_window.print_control_identifiers)

# Save to log file
log_file = "mseq_controls.log"
with open(log_file, 'w') as f:
    f.write(output)

print(f"Control identifiers saved to {log_file}")

# Also try to capture all windows/dialogs
print("Capturing information about all windows in the application...")
windows_output = "All windows in the application:\n\n"
for window in app.windows():
    windows_output += f"Window: {window.window_text()}\n"
    try:
        window_output = capture_print_output(window.print_control_identifiers)
        windows_output += window_output + "\n" + "-"*80 + "\n\n"
    except:
        windows_output += "Failed to get control identifiers for this window\n" + "-"*80 + "\n\n"

with open("mseq_all_windows.log", 'w') as f:
    f.write(windows_output)

print(f"All windows information saved to mseq_all_windows.log")
print("Script completed successfully.")