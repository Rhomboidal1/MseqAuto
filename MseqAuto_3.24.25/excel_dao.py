# excel_dao.py
import os
import pylightxl as xl
from datetime import datetime

class ExcelDAO:
    def __init__(self, config):
        self.config = config
    
    def create_workbook(self):
        """Create a new workbook"""
        db = xl.Database()
        db.add_ws(ws="Sheet1")
        return db
    
    def load_workbook(self, file_path):
        """Load an existing workbook"""
        if not os.path.exists(file_path):
            return None
        try:
            return xl.readxl(file_path)
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return None
    
    def save_workbook(self, workbook, file_path):
        """Save workbook to file"""
        try:
            xl.writexl(db=workbook, fn=file_path)
            return file_path
        except Exception as e:
            print(f"Error saving Excel file: {e}")
            return None
    
    def set_cell_value(self, worksheet, row, col, value):
        """Set cell value"""
        # pylightxl uses 1-based indexing like Excel
        worksheet.update_index(row=row, col=col, val=value)
    
    def set_header_row(self, worksheet, headers):
        """Set header row values"""
        for i, header in enumerate(headers, 1):
            self.set_cell_value(worksheet, 1, i, header)
    
    def get_cell_value(self, worksheet, row, col):
        """Get cell value"""
        # pylightxl requires using index-based access
        return worksheet.index(row, col)
    
    def get_max_row(self, worksheet):
        """Get maximum row number"""
        return worksheet.maxrow
    
    def create_validation_summary(self, worksheet, data_folder, prefix=None):
        """Create a validation summary workbook"""
        # Set headers
        headers = ['I Number', 'Order Number', 'Status', 'Zip Filename', 
                  'Order Items', 'File Names', 'Match Status', 'Zip Timestamp']
        self.set_header_row(worksheet, headers)
        
        return worksheet
    
    def add_validation_result(self, worksheet, row_count, validation_result, zip_path, i_number, order_number, is_andreev=False):
        """Add validation result to worksheet"""
        # Set basic information
        self.set_cell_value(worksheet, row_count, 1, i_number)
        self.set_cell_value(worksheet, row_count, 2, order_number)
        self.set_cell_value(worksheet, row_count, 4, os.path.basename(zip_path))
        self.set_cell_value(worksheet, row_count, 8, str(int(os.path.getmtime(zip_path))))
        
        order_row = row_count
        row_count += 1
        
        # Set status based on validation results
        match_count = validation_result.get('match_count', 0)
        expected_count = validation_result.get('expected_count', 0)
        txt_count = validation_result.get('txt_count', 0)
        
        if is_andreev:
            # Andreev special case - only care about AB1 files, not TXT
            if match_count == expected_count:
                self.set_cell_value(worksheet, order_row, 3, 'Completed')
            else:
                self.set_cell_value(worksheet, order_row, 3, 'ATTENTION')
        else:
            # Normal case - need all AB1 and TXT files
            if match_count == expected_count and txt_count == 5:
                self.set_cell_value(worksheet, order_row, 3, 'Completed')
            else:
                self.set_cell_value(worksheet, order_row, 3, 'ATTENTION')
        
        # Add matches
        for match in validation_result.get('matches', []):
            self.set_cell_value(worksheet, row_count, 5, match['raw_name'])
            self.set_cell_value(worksheet, row_count, 6, match['file_name'])
            self.set_cell_value(worksheet, row_count, 7, 'match')
            row_count += 1
        
        # Add mismatches in zip
        for mismatch in validation_result.get('mismatches_in_zip', []):
            self.set_cell_value(worksheet, row_count, 6, mismatch)
            self.set_cell_value(worksheet, row_count, 7, 'no match')
            row_count += 1
        
        # Add mismatches in order
        for mismatch in validation_result.get('mismatches_in_order', []):
            self.set_cell_value(worksheet, row_count, 5, mismatch['raw_name'])
            self.set_cell_value(worksheet, row_count, 7, 'no match')
            row_count += 1
        
        # Add txt files
        for txt_ext in self.config.TEXT_FILES:
            self.set_cell_value(worksheet, row_count, 5, txt_ext)
            if txt_ext in validation_result.get('txt_files', []):
                self.set_cell_value(worksheet, row_count, 6, f"*{txt_ext}")
                self.set_cell_value(worksheet, row_count, 7, 'txt file')
            else:
                self.set_cell_value(worksheet, row_count, 7, 'MISSING txt file')
            row_count += 1
        
        return row_count


if __name__ == "__main__":
    # Simple test if run directly
    from config import MseqConfig
    
    class TestConfig:
        TEXT_FILES = ['.raw.qual.txt', '.raw.seq.txt', '.seq.info.txt', '.seq.qual.txt', '.seq.txt']
    
    config = TestConfig()
    excel_dao = ExcelDAO(config)
    
    # Create a test workbook
    wb = excel_dao.create_workbook()
    ws = wb.ws("Sheet1")
    
    # Set some test headers
    excel_dao.set_header_row(ws, ["Column A", "Column B", "Column C"])
    
    # Set some test values
    excel_dao.set_cell_value(ws, 2, 1, "Test Value A2")
    excel_dao.set_cell_value(ws, 2, 2, "Test Value B2")
    excel_dao.set_cell_value(ws, 2, 3, "Test Value C2")
    
    # Print test outputs
    print(f"Max row: {excel_dao.get_max_row(ws)}")
    print(f"Cell B2: {excel_dao.get_cell_value(ws, 2, 2)}")
    
    # Optionally save the workbook
    # excel_dao.save_workbook(wb, "test_workbook.xlsx")