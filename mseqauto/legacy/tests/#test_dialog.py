# test_dialog.py
import tkinter as tk
from tkinter import filedialog
import time

def get_folder_from_user():
    """Get folder selection from user using a simple approach"""
    print("Opening folder selection dialog...")

    # Create and immediately withdraw (hide) the root window
    root = tk.Tk()
    root.withdraw()

    # Show a directory selection dialog
    folder_path = filedialog.askdirectory(
        title="Select today's data folder to mseq orders",
        mustexist=True
    )

    # Destroy the root window
    root.destroy()

    if folder_path:
        print(f"Selected folder: {folder_path}")
        return folder_path
    else:
        print("No folder selected")
        return None

print("Starting test script")
print("About to show dialog")

# Method 1 - direct call
result1 = filedialog.askdirectory(title="Test Dialog 1")
print(f"Result 1: {result1}")

time.sleep(1)

# Method 2 - with root
root = tk.Tk()
root.withdraw()
result2 = filedialog.askdirectory(title="Test Dialog 2")
root.destroy()
print(f"Result 2: {result2}")

print("Test completed")

get_folder_from_user()