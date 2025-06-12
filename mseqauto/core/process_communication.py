"""Simple communication protocol for GUI-script interaction"""
import sys
import json
import os

# Check if running under GUI
GUI_MODE = os.environ.get('MSEQAUTO_GUI_MODE', 'False') == 'True'

def send_gui_message(msg_type, data):
    """Send message to GUI if in GUI mode"""
    if GUI_MODE:
        message = {"type": msg_type, "data": data}
        print(f"::MSG::{json.dumps(message)}::ENDMSG::", flush=True)

def send_progress(current, total, description=""):
    send_gui_message("progress", {
        "current": current, 
        "total": total, 
        "description": description
    })

def send_status(status):
    send_gui_message("status", {"status": status})

# Make these available but do nothing in standalone mode
if not GUI_MODE:
    send_progress = lambda *args, **kwargs: None
    send_status = lambda *args, **kwargs: None