# test_tkinter_dialog.py
import tkinter as tk
from tkinter import filedialog
import os

def main():
    print("Testing Tkinter folder dialog...")
    
    # Create and hide the main Tkinter window
    root = tk.Tk()
    root.withdraw()
    
    # Create a smaller window to show we're running
    dialog_window = tk.Toplevel(root)
    dialog_window.title("Folder Dialog Test")
    dialog_window.geometry("300x200")
    
    # Add a label and buttons
    label = tk.Label(dialog_window, text="Click 'Select Folder' to open the dialog")
    label.pack(pady=20)
    
    result_label = tk.Label(dialog_window, text="No folder selected yet")
    result_label.pack(pady=10)
    
    def on_select_folder():
        # Open the folder selection dialog
        folder_path = filedialog.askdirectory(title="Select a folder for testing")
        
        if folder_path:
            # Update the result label with the selected path
            result_label.config(text=f"Selected: {folder_path}")
            print(f"Selected folder: {folder_path}")
        else:
            # User cancelled
            result_label.config(text="Selection cancelled")
            print("Selection cancelled")
    
    # Create a button to open the dialog
    select_button = tk.Button(dialog_window, text="Select Folder", command=on_select_folder)
    select_button.pack(pady=10)
    
    # Create a quit button
    quit_button = tk.Button(dialog_window, text="Quit", command=root.destroy)
    quit_button.pack(pady=10)
    
    # Start the Tkinter event loop
    root.mainloop()
    
    print("Tkinter test completed")

if __name__ == "__main__":
    main()