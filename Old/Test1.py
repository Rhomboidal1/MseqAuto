from tkinter import filedialog
from os import listdir, path as OsPath
from re import sub
from time import sleep

# Import the necessary pywinauto modules
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
from pywinauto.keyboard import send_keys
from pywinauto import timings

def GetMseq():
    print("Attempting to connect to or start mSeq...")
    try:
        app = Application(backend='win32').connect(title_re='Mseq*', timeout=1)
        print("Connected to existing 'Mseq*' window")
    except (ElementNotFoundError, timings.TimeoutError):
        try:
            app = Application(backend='win32').connect(title_re='mSeq*', timeout=1)
            print("Connected to existing 'mSeq*' window")
        except (ElementNotFoundError, timings.TimeoutError):
            # This is the EXACT line from the annotated version - the key is to chain these methods
            print("Starting new mSeq instance...")
            app = Application(backend='win32').start('cmd /c "cd /d C:\\DNA\\Mseq4\\bin && C:\\DNA\\Mseq4\\bin\\j.exe -jprofile mseq.ijl"', wait_for_idle=False).connect(title='mSeq', timeout=10)
            print("Started and connected to new mSeq instance")
        except ElementAmbiguousError:
            app = Application(backend='win32').connect(title_re='mSeq*', found_index=0, timeout=1)
            app2 = Application(backend='win32').connect(title_re='mSeq*', found_index=1, timeout=1)
            app2.kill()
            print("Handled multiple 'mSeq*' windows")
    except ElementAmbiguousError:
        app = Application(backend='win32').connect(title_re='Mseq*', found_index=0, timeout=1)
        app2 = Application(backend='win32').connect(title_re='Mseq*', found_index=1, timeout=1)
        app2.kill()
        print("Handled multiple 'Mseq*' windows")
    
    if app.window(title_re='mSeq*').exists()==False:
        mainWindow = app.window(title_re='Mseq*')
        print("Using 'Mseq*' window")
    else:
        mainWindow = app.window(title_re='mSeq*')
        print("Using 'mSeq*' window")

    print("mSeq found successfully")
    return app, mainWindow

# Main program
if __name__ == "__main__":
    print("=== Simplified mSeq Launcher ===")
    print("Trying to start mSeq...")
    
    try:
        app, mainWindow = GetMseq()
        print("Successfully got mSeq application and main window")
        print(f"Main window title: {mainWindow.window_text()}")
        
        # Focus the window to make it visible
        mainWindow.set_focus()
        print("Window focused - you should see the mSeq application on screen now")
        
        # Wait before exiting
        print("Script completed successfully. Window will remain open.")
        input("Press Enter to exit...")
        
    except Exception as e:
        print(f"An error occurred: {e}")
        input("Press Enter to exit...")