"""MseqAuto GUI Application"""
import sys
import os
import subprocess
import json
from datetime import datetime
from pathlib import Path
import warnings

from PyQt6.QtWidgets import * #type: ignore
from PyQt6.QtCore import * #type: ignore
from PyQt6.QtGui import * #type: ignore


sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", module="pywinauto")

from mseqauto.config import MseqConfig # type: ignore

class WorkerThread(QThread):
    """Simple thread to run 32-bit Python scripts"""
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, int, str)
    status_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, script_name, data_folder):
        super().__init__()
        self.script_name = script_name
        self.data_folder = data_folder
        self.config = MseqConfig()
        self.process = None
        
    def run(self):
        """Run the script in 32-bit Python"""
        try:
            # Build command
            script_path = os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                'scripts',
                self.script_name
            )
            
            # Set up environment
            env = os.environ.copy()
            env['MSEQAUTO_GUI_MODE'] = 'True'
            env['MSEQAUTO_DATA_FOLDER'] = self.data_folder
            
            # Use 32-bit Python
            cmd = [self.config.PYTHON32_PATH, script_path]
            
            # Create process
            self.process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                env=env
            )
            
            # Read output
            for line in iter(self.process.stdout.readline, ''): #type: ignore
                if line:
                    self.process_line(line.strip())
                    
            # Wait for completion
            exit_code = self.process.wait()
            self.finished_signal.emit(exit_code == 0, 
                                    f"Process exited with code {exit_code}")
            
        except Exception as e:
            self.finished_signal.emit(False, str(e))
            
    def process_line(self, line):
        """Process output line"""
        # Check for structured message
        if line.startswith("::MSG::") and line.endswith("::ENDMSG::"):
            try:
                json_str = line[7:-9]
                message = json.loads(json_str)
                self.handle_message(message)
                return
            except:
                pass
        
        # Regular log line
        self.log_signal.emit(line)
        
    def handle_message(self, message):
        """Handle structured message"""
        msg_type = message.get('type')
        data = message.get('data', {})
        
        if msg_type == 'progress':
            self.progress_signal.emit(
                data.get('current', 0),
                data.get('total', 100),
                data.get('description', '')
            )
        elif msg_type == 'status':
            self.status_signal.emit(data.get('status', ''))

# Add this to your gui/main.py

class LogViewerWidget(QWidget):
    """Widget for viewing historical log files"""
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # Controls
        controls_layout = QHBoxLayout()
        
        controls_layout.addWidget(QLabel("Log Files:"))
        
        self.log_combo = QComboBox()
        self.log_combo.currentTextChanged.connect(self.load_selected_log)
        controls_layout.addWidget(self.log_combo, 1)
        
        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.clicked.connect(self.refresh_log_list)
        controls_layout.addWidget(self.refresh_btn)
        
        self.delete_btn = QPushButton("Delete")
        self.delete_btn.clicked.connect(self.delete_selected_log)
        controls_layout.addWidget(self.delete_btn)
        
        layout.addLayout(controls_layout)
        
        # Search bar
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Search:"))
        self.search_input = QLineEdit()
        self.search_input.textChanged.connect(self.search_logs)
        search_layout.addWidget(self.search_input)
        
        self.case_sensitive_cb = QCheckBox("Case Sensitive")
        search_layout.addWidget(self.case_sensitive_cb)
        
        layout.addLayout(search_layout)
        
        # Log viewer
        self.log_viewer = QTextEdit()
        self.log_viewer.setReadOnly(True)
        self.log_viewer.setFont(QFont("Consolas", 9))
        layout.addWidget(self.log_viewer)
        
        # Initial load
        self.refresh_log_list()
        
    def get_log_directory(self):
        """Get the log directory path"""
        return os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'scripts', 'logs'
        )
        
    def refresh_log_list(self):
        """Refresh the list of log files"""
        log_dir = self.get_log_directory()
        
        if not os.path.exists(log_dir):
            return
            
        # Get all log files
        log_files = []
        for file in os.listdir(log_dir):
            if file.endswith('.log'):
                # Get file stats for sorting
                file_path = os.path.join(log_dir, file)
                mtime = os.path.getmtime(file_path)
                log_files.append((file, mtime))
                
        # Sort by modification time (newest first)
        log_files.sort(key=lambda x: x[1], reverse=True)
        
        # Update combo box
        self.log_combo.clear()
        self.log_combo.addItem("-- Select a log file --")
        for file, _ in log_files:
            self.log_combo.addItem(file)
            
    def load_selected_log(self, filename):
        """Load the selected log file"""
        if not filename or filename.startswith("--"):
            self.log_viewer.clear()
            return
            
        log_path = os.path.join(self.get_log_directory(), filename)
        
        try:
            with open(log_path, 'r', encoding='utf-8') as f:
                content = f.read()
                self.log_viewer.setPlainText(content)
                
            # Get file info
            stat = os.stat(log_path)
            size_kb = stat.st_size / 1024
            mod_time = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            
            # Add file info at the top
            info = f"File: {filename}\nSize: {size_kb:.1f} KB\nModified: {mod_time}\n{'='*60}\n\n"
            self.log_viewer.setPlainText(info + content)
            
        except Exception as e:
            self.log_viewer.setPlainText(f"Error loading log file: {str(e)}")
            
    def search_logs(self):
        """Search within the current log"""
        search_text = self.search_input.text()
        if not search_text:
            # Clear any existing highlighting
            cursor = self.log_viewer.textCursor()
            cursor.select(QTextCursor.SelectionType.Document)
            cursor.setCharFormat(QTextCharFormat())
            cursor.clearSelection()
            self.log_viewer.setTextCursor(cursor)
            return
            
        # Search flags
        flags = QTextDocument.FindFlag(0)
        if self.case_sensitive_cb.isChecked():
            flags |= QTextDocument.FindFlag.FindCaseSensitively
            
        # Clear previous highlighting
        cursor = self.log_viewer.textCursor()
        cursor.select(QTextCursor.SelectionType.Document)
        cursor.setCharFormat(QTextCharFormat())
        cursor.clearSelection()
        self.log_viewer.setTextCursor(cursor)
        
        # Highlight all occurrences
        highlight_format = QTextCharFormat()
        highlight_format.setBackground(QColor("yellow"))
        
        cursor = self.log_viewer.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.Start)
        
        found = False
        while True:
            cursor = self.log_viewer.document().find(search_text, cursor, flags) #type: ignore
            if cursor.isNull():
                break
                
            found = True
            cursor.mergeCharFormat(highlight_format)
            
        # Move to first occurrence
        if found:
            cursor = self.log_viewer.textCursor()
            cursor.movePosition(QTextCursor.MoveOperation.Start)
            first_cursor = self.log_viewer.document().find(search_text, cursor, flags) #type: ignore
            if not first_cursor.isNull():
                self.log_viewer.setTextCursor(first_cursor)
                
    def delete_selected_log(self):
        """Delete the selected log file"""
        filename = self.log_combo.currentText()
        if not filename or filename.startswith("--"):
            return
            
        reply = QMessageBox.question(
            self, 
            "Delete Log File",
            f"Are you sure you want to delete {filename}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            log_path = os.path.join(self.get_log_directory(), filename)
            try:
                os.remove(log_path)
                self.refresh_log_list()
                self.log_viewer.clear()
                QMessageBox.information(self, "Success", "Log file deleted")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete log: {str(e)}")

# Modified MseqAutoGUI class - add the log viewer tab
class MseqAutoGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.data_folder = None
        self.worker = None
        self.init_ui()
        
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("DNA Sequencing Automation")
        self.setGeometry(100, 100, 1000, 700)
        
        # Central widget with tabs
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        
        # Folder selection (above tabs)
        folder_layout = QHBoxLayout()
        folder_layout.addWidget(QLabel("Data Folder:"))
        self.folder_label = QLabel("No folder selected")
        folder_layout.addWidget(self.folder_label, 1)
        self.select_folder_btn = QPushButton("Select Folder")
        self.select_folder_btn.clicked.connect(self.select_folder)
        folder_layout.addWidget(self.select_folder_btn)
        layout.addLayout(folder_layout)
        
        # Tab widget
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)
        
        # Workflow tab
        workflow_widget = self.create_workflow_tab()
        self.tabs.addTab(workflow_widget, "Workflows")
        
        # Log viewer tab
        self.log_viewer = LogViewerWidget()
        self.tabs.addTab(self.log_viewer, "Log History")
        
        # Status bar
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Ready") #type: ignore
        
        # Initial state
        self.update_button_states()
        
    def create_workflow_tab(self):
        """Create the workflow tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Workflow buttons
        workflow_group = QGroupBox("Workflows")
        workflow_layout = QGridLayout()
        
        # Individual workflows
        ind_label = QLabel("<b>Individual Sequencing</b>")
        workflow_layout.addWidget(ind_label, 0, 0, 1, 4)
        
        self.ind_sort_btn = QPushButton("Sort Files")
        self.ind_sort_btn.clicked.connect(lambda: self.run_script('ind_sort_files.py'))
        workflow_layout.addWidget(self.ind_sort_btn, 1, 0)
        
        self.ind_mseq_btn = QPushButton("Run mSeq")
        self.ind_mseq_btn.clicked.connect(lambda: self.run_script('ind_auto_mseq.py'))
        workflow_layout.addWidget(self.ind_mseq_btn, 1, 1)
        
        self.ind_zip_btn = QPushButton("Zip Files")
        self.ind_zip_btn.clicked.connect(lambda: self.run_script('ind_zip_files.py'))
        workflow_layout.addWidget(self.ind_zip_btn, 1, 2)
        
        self.ind_all_btn = QPushButton("Run All")
        self.ind_all_btn.clicked.connect(lambda: self.run_script('ind_process_all.py'))
        workflow_layout.addWidget(self.ind_all_btn, 1, 3)
        
        # Plate workflows
        plate_label = QLabel("<b>Plate Sequencing</b>")
        workflow_layout.addWidget(plate_label, 2, 0, 1, 4)
        
        self.plate_sort_btn = QPushButton("Sort Complete")
        self.plate_sort_btn.clicked.connect(lambda: self.run_script('plate_sort_complete.py'))
        workflow_layout.addWidget(self.plate_sort_btn, 3, 0)
        
        self.plate_mseq_btn = QPushButton("Run mSeq")
        self.plate_mseq_btn.clicked.connect(lambda: self.run_script('plate_auto_mseq.py'))
        workflow_layout.addWidget(self.plate_mseq_btn, 3, 1)
        
        self.plate_zip_btn = QPushButton("Zip Files")
        self.plate_zip_btn.clicked.connect(lambda: self.run_script('plate_zip_files.py'))
        workflow_layout.addWidget(self.plate_zip_btn, 3, 2)
        
        # Add Full Plasmid button
        plasmid_label = QLabel("<b>Special Workflows</b>")
        workflow_layout.addWidget(plasmid_label, 4, 0, 1, 4)
        
        self.plasmid_zip_btn = QPushButton("Zip Full Plasmid Orders")
        self.plasmid_zip_btn.clicked.connect(lambda: self.run_script('full_plasmid_zip_files.py'))
        workflow_layout.addWidget(self.plasmid_zip_btn, 5, 0)
        
        workflow_group.setLayout(workflow_layout)
        layout.addWidget(workflow_group)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        
        # Console output
        console_group = QGroupBox("Console Output")
        console_layout = QVBoxLayout()
        
        self.console_text = QTextEdit()
        self.console_text.setReadOnly(True)
        self.console_text.setFont(QFont("Consolas", 9))
        console_layout.addWidget(self.console_text)
        
        # Console buttons
        console_btn_layout = QHBoxLayout()
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self.console_text.clear)
        console_btn_layout.addWidget(clear_btn)
        
        save_btn = QPushButton("Save Current Output")
        save_btn.clicked.connect(self.save_current_output)
        console_btn_layout.addWidget(save_btn)
        console_btn_layout.addStretch()
        
        console_layout.addLayout(console_btn_layout)
        console_group.setLayout(console_layout)
        layout.addWidget(console_group)
        
        return widget
        
    def on_worker_finished(self, success, message): #type: ignore
        """Handle worker completion - also refresh log list"""
        self.worker = None
        self.progress_bar.setValue(0)
        self.update_button_states()
        
        if success:
            self.log(f"\n✓ Completed successfully")
            self.status_bar.showMessage("Ready") #type: ignore
            # Refresh log list to show new log
            self.log_viewer.refresh_log_list()
        else:
            self.log(f"\n✗ Failed: {message}")
            self.status_bar.showMessage(f"Error: {message}") #type: ignore
            
            
    def save_current_output(self):
        """Save current console output"""
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Save Current Output",
            f"mseqauto_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            "Text Files (*.txt)"
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.console_text.toPlainText())
                self.log(f"Output saved to: {filename}")
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Failed to save output: {str(e)}")
        
    def select_folder(self):
        """Select data folder"""
        folder = QFileDialog.getExistingDirectory(
            self, 
            "Select Data Folder",
            self.data_folder or ""
        )
        if folder:
            self.data_folder = folder
            self.folder_label.setText(folder)
            self.update_button_states()
            self.log(f"Selected folder: {folder}")
            
    def update_button_states(self):
        """Enable/disable buttons based on state"""
        has_folder = self.data_folder is not None
        is_running = self.worker is not None
        
        # Workflow buttons
        for btn in [self.ind_sort_btn, self.ind_mseq_btn, self.ind_zip_btn,
                   self.ind_all_btn, self.plate_sort_btn, self.plate_mseq_btn,
                   self.plate_zip_btn]:
            btn.setEnabled(has_folder and not is_running)
            
        # Folder selection
        self.select_folder_btn.setEnabled(not is_running)
        
    def run_script(self, script_name):
        """Run a script in 32-bit Python"""
        if not self.data_folder:
            QMessageBox.warning(self, "No Folder", "Please select a data folder first")
            return
            
        self.log(f"\n{'='*60}")
        self.log(f"Starting: {script_name}")
        self.log(f"{'='*60}")
        
        # Create and start worker
        self.worker = WorkerThread(script_name, self.data_folder)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.status_signal.connect(self.status_bar.showMessage)  #type: ignore
        self.worker.finished_signal.connect(self.on_worker_finished)
        
        self.worker.start()
        self.update_button_states()
        
    def log(self, message):
        """Add message to console"""
        self.console_text.append(message)
        
    def update_progress(self, current, total, description):
        """Update progress bar"""
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        if description:
            self.status_bar.showMessage(description)  #type: ignore
            

    def save_log(self):
        """Save console log to file"""
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Save Log",
            f"mseqauto_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            "Text Files (*.txt)"
        )
        if filename:
            with open(filename, 'w') as f:
                f.write(self.console_text.toPlainText())
            self.log(f"Log saved to: {filename}")

def main():
    app = QApplication(sys.argv)
    window = MseqAutoGUI()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()