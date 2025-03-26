# excel_dao.py
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


class ExcelDAO:
    def __init__(self, config):
        self.config = config
        # Define standard styles
        self.success_style = PatternFill(start_color='00CC00', end_color='00CC00', fill_type='solid')
        self.attention_style = PatternFill(start_color='FF4747', end_color='FF4747', fill_type='solid')
        self.resolved_style = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
        self.break_style = PatternFill(start_color='DDD9C4', end_color='DDD9C4', fill_type='solid')

    def create_workbook(self):
        """Create a new workbook"""
        return Workbook()

    def load_workbook(self, file_path):
        """Load an existing workbook"""
        if not os.path.exists(file_path):
            return None
        try:
            return load_workbook(file_path)
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return None

    def save_workbook(self, workbook, file_path):
        """Save workbook to file"""
        try:
            workbook.save(file_path)
            return file_path
        except Exception as e:
            print(f"Error saving Excel file: {e}")
            return None

    def set_cell_value(self, worksheet, row, col, value):
        """Set cell value"""
        worksheet.cell(row=row, column=col, value=value)

    def set_header_row(self, worksheet, headers):
        """Set header row values"""
        for i, header in enumerate(headers, 1):
            cell = worksheet.cell(row=1, column=i, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

    def get_cell_value(self, worksheet, row, col):
        """Get cell value"""
        return worksheet.cell(row=row, column=col).value

    def get_max_row(self, worksheet):
        """Get maximum row number"""
        return worksheet.max_row

    def apply_style(self, worksheet, cell_ref, style_type):
        """Apply style to cell

        Args:
            worksheet: The worksheet object
            cell_ref: Cell reference (e.g., 'A1')
            style_type: Type of style to apply ('success', 'attention', 'break', 'resolved')
        """
        if style_type == 'success':
            worksheet[cell_ref].fill = self.success_style
        elif style_type == 'attention':
            worksheet[cell_ref].fill = self.attention_style
        elif style_type == 'break':
            worksheet[cell_ref].fill = self.break_style
        elif style_type == 'resolved':
            worksheet[cell_ref].fill = self.resolved_style

    def adjust_column_widths(self, worksheet):
        """Adjust column widths based on content"""
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)

            for cell in column:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length

            adjusted_width = max_length + 2
            worksheet.column_dimensions[column_letter].width = adjusted_width

    def create_validation_summary(self, worksheet, _data_folder, _prefix=None):
        """Create a validation summary workbook"""
        # Set headers
        headers = ['I Number', 'Order Number', 'Status', 'Zip Filename',
                   'Order Items', 'File Names', 'Match Status', 'Zip Timestamp']
        self.set_header_row(worksheet, headers)

        return worksheet

    def add_validation_result(self, worksheet, row_count, validation_result, zip_path, i_number, order_number,
                              is_andreev=False):
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
                self.apply_style(worksheet, f'C{order_row}', 'success')
            else:
                self.set_cell_value(worksheet, order_row, 3, 'ATTENTION')
                self.apply_style(worksheet, f'C{order_row}', 'attention')
        else:
            # Normal case - need all AB1 and TXT files
            if match_count == expected_count and txt_count == 5:
                self.set_cell_value(worksheet, order_row, 3, 'Completed')
                self.apply_style(worksheet, f'C{order_row}', 'success')
            else:
                self.set_cell_value(worksheet, order_row, 3, 'ATTENTION')
                self.apply_style(worksheet, f'C{order_row}', 'attention')

        # Add matches
        for match in validation_result.get('matches', []):
            self.set_cell_value(worksheet, row_count, 5, match['raw_name'])
            self.set_cell_value(worksheet, row_count, 6, match['file_name'])
            self.set_cell_value(worksheet, row_count, 7, 'match')
            self.apply_style(worksheet, f'G{row_count}', 'success')
            row_count += 1

        # Add mismatches in zip
        for mismatch in validation_result.get('mismatches_in_zip', []):
            self.set_cell_value(worksheet, row_count, 6, mismatch)
            self.set_cell_value(worksheet, row_count, 7, 'no match')
            self.apply_style(worksheet, f'G{row_count}', 'attention')
            row_count += 1

        # Add mismatches in order
        for mismatch in validation_result.get('mismatches_in_order', []):
            self.set_cell_value(worksheet, row_count, 5, mismatch['raw_name'])
            self.set_cell_value(worksheet, row_count, 7, 'no match')
            self.apply_style(worksheet, f'G{row_count}', 'attention')
            row_count += 1

        # Add txt files
        for txt_ext in self.config.TEXT_FILES:
            self.set_cell_value(worksheet, row_count, 5, txt_ext)
            if txt_ext in validation_result.get('txt_files', []):
                self.set_cell_value(worksheet, row_count, 6, f"*{txt_ext}")
                self.set_cell_value(worksheet, row_count, 7, 'txt file')
                self.apply_style(worksheet, f'G{row_count}', 'success')
            else:
                self.set_cell_value(worksheet, row_count, 7, 'MISSING txt file')
                self.apply_style(worksheet, f'G{row_count}', 'attention')
            row_count += 1

        # Hide rows for completed orders if needed
        if self.get_cell_value(worksheet, order_row, 3) == 'Completed':
            for i in range(order_row + 1, row_count):
                worksheet.row_dimensions[i].hidden = True

        return row_count


if __name__ == "__main__":
    # Simple test if run directly
    class TestConfig:
        TEXT_FILES = ['.raw.qual.txt', '.raw.seq.txt', '.seq.info.txt', '.seq.qual.txt', '.seq.txt']


    config = TestConfig()
    excel_dao = ExcelDAO(config)

    # Create a test workbook
    wb = excel_dao.create_workbook()
    ws = wb.active

    # Set some test headers
    excel_dao.set_header_row(ws, ["Column A", "Column B", "Column C"])

    # Set some test values
    excel_dao.set_cell_value(ws, 2, 1, "Test Value A2")
    excel_dao.set_cell_value(ws, 2, 2, "Test Value B2")
    excel_dao.set_cell_value(ws, 2, 3, "Test Value C2")

    # Apply test styles
    excel_dao.apply_style(ws, "A2", "success")
    excel_dao.apply_style(ws, "B2", "attention")
    excel_dao.apply_style(ws, "C2", "break")

    # Print test outputs
    print(f"Max row: {excel_dao.get_max_row(ws)}")
    print(f"Cell B2: {excel_dao.get_cell_value(ws, 2, 2)}")

    # Optionally save the workbook
    # excel_dao.save_workbook(wb, "test_workbook.xlsx")