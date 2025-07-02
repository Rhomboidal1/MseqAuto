# excel_dao.py - Simplified version
import os
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


class ExcelDAO:
    def __init__(self, config):
        self.config = config
        # Define standard styles - keeping exact same colors as original
        self.success_style = PatternFill(start_color='00CC00', end_color='00CC00', fill_type='solid')
        self.attention_style = PatternFill(start_color='FF4747', end_color='FF4747', fill_type='solid')
        self.resolved_style = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
        self.break_style = PatternFill(start_color='DDD9C4', end_color='DDD9C4', fill_type='solid')
        
        # Store hidden row states for preservation during updates
        self._stored_hidden_states = {}

    # Core workbook operations
    def create_workbook(self):
        """Create a new workbook"""
        return Workbook()

    def load_workbook(self, file_path):
        """Load an existing workbook with error handling"""
        if not Path(file_path).exists():
            return None
        try:
            return load_workbook(file_path)
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return None

    def save_with_error_handling(self, workbook, file_path):
        """Save workbook with specific error handling for permission issues"""
        try:
            workbook.save(file_path)
            return True
        except PermissionError:
            print(f"Error: Cannot save Excel file. Please close {Path(file_path).name} in Excel and try again.")
            return False
        except Exception as e:
            print(f"Error saving Excel file: {e}")
            return False

    # Cell operations
    def set_cell_value(self, worksheet, row, col, value):
        """Set cell value"""
        worksheet.cell(row=row, column=col, value=value)

    def get_cell_value(self, worksheet, row, col):
        """Get cell value"""
        return worksheet.cell(row=row, column=col).value

    def apply_style(self, worksheet, cell_ref, style_type):
        """Apply style to cell"""
        style_map = {
            'success': self.success_style,
            'attention': self.attention_style,
            'break': self.break_style,
            'resolved': self.resolved_style
        }
        if style_type in style_map:
            worksheet[cell_ref].fill = style_map[style_type]

    # Formatting operations
    def set_validation_headers(self, worksheet):
        """Set headers for validation summary"""
        headers = ['I Number', 'Order Number', 'Status', 'Zip Filename',
                   'Order Items', 'File Names', 'Match Status', 'Zip Timestamp']
        for i, header in enumerate(headers, 1):
            cell = worksheet.cell(row=1, column=i, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

    def auto_adjust_columns(self, worksheet):
        """Adjust column widths based on content - matches original script logic"""
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)

            for cell in column:
                if cell.value:
                    try:
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                    except:
                        pass

            adjusted_width = max_length + 2
            worksheet.column_dimensions[column_letter].width = adjusted_width

    # Row management for updates
    def preserve_hidden_row_states(self, worksheet):
        """Store the current hidden state of all rows"""
        self._stored_hidden_states = {}
        for i in range(2, worksheet.max_row + 1):  # Start from 2 to skip header
            self._stored_hidden_states[i] = worksheet.row_dimensions[i].hidden

    def restore_hidden_row_states(self, worksheet, offset=0):
        """Restore hidden states with optional offset for inserted rows"""
        for original_row, hidden_state in self._stored_hidden_states.items():
            new_row = original_row + offset
            if new_row <= worksheet.max_row:
                worksheet.row_dimensions[new_row].hidden = hidden_state

    def insert_rows_at_top(self, worksheet, num_rows):
        """Insert rows at top (after header) and handle hidden state preservation"""
        if num_rows <= 0:
            return
        self.preserve_hidden_row_states(worksheet)
        worksheet.insert_rows(2, num_rows)

    # Order management
    def find_order_in_summary(self, worksheet, order_number):
        """Find if an order already exists in the summary"""
        for row_num, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
            if str(row[1]) == str(order_number):  # Column B is order number
                zip_timestamp = row[7] if len(row) > 7 else None  # Column H is zip timestamp
                return True, row_num, zip_timestamp
        return False, None, None

    def resolve_order_status(self, worksheet, order_number):
        """Mark an order as resolved and hide associated rows"""
        found, order_row, _ = self.find_order_in_summary(worksheet, order_number)
        
        if not found:
            return False
            
        # Set status to "Resolved"
        self.set_cell_value(worksheet, order_row, 3, "Resolved")  # Column C
        self.apply_style(worksheet, f'C{order_row}', 'resolved')
        
        # Hide rows until we hit another order or end of data
        current_row = order_row + 1 #type: ignore
        while current_row <= worksheet.max_row:
            order_cell_value = self.get_cell_value(worksheet, current_row, 2)
            if order_cell_value is not None and str(order_cell_value).strip():
                break  # Hit another order, stop hiding
            else:
                worksheet.row_dimensions[current_row].hidden = True
                current_row += 1
                
        return True

    def add_break_row(self, worksheet, row_num):
        """Add a break row with styling"""
        self.set_cell_value(worksheet, row_num, 1, "Break")
        self.apply_style(worksheet, f'A{row_num}', 'break')
        worksheet.row_dimensions[row_num].hidden = False

    # Main validation result processing
    def add_validation_result(self, worksheet, row_count, validation_result, zip_path, i_number, order_number,
                              is_andreev=False):
        """Add validation result to worksheet"""
        # Set basic information
        self.set_cell_value(worksheet, row_count, 1, i_number)
        self.set_cell_value(worksheet, row_count, 2, order_number)
        self.set_cell_value(worksheet, row_count, 4, Path(zip_path).name)
        self.set_cell_value(worksheet, row_count, 8, str(int(Path(zip_path).stat().st_mtime)))

        order_row = row_count
        row_count += 1

        # Determine status
        match_count = validation_result.get('match_count', 0)
        expected_count = validation_result.get('expected_count', 0)
        txt_count = validation_result.get('txt_count', 0)
        extra_ab1_count = validation_result.get('extra_ab1_count', 0)

        # Status logic: completed if all expected files match, no extra files, and (for non-Andreev) all txt files present
        is_completed = (match_count == expected_count and match_count != 0 and extra_ab1_count == 0 and
                       (is_andreev or txt_count == 5))

        status = 'Completed' if is_completed else 'ATTENTION'
        style = 'success' if is_completed else 'attention'
        
        self.set_cell_value(worksheet, order_row, 3, status)
        self.apply_style(worksheet, f'C{order_row}', style)

        # Add file details
        for match in validation_result.get('matches', []):
            self.set_cell_value(worksheet, row_count, 5, match['raw_name'])
            self.set_cell_value(worksheet, row_count, 6, match['file_name'])
            self.set_cell_value(worksheet, row_count, 7, 'match')
            self.apply_style(worksheet, f'G{row_count}', 'success')
            row_count += 1

        for mismatch in validation_result.get('mismatches_in_zip', []):
            self.set_cell_value(worksheet, row_count, 6, mismatch)
            self.set_cell_value(worksheet, row_count, 7, 'no match')
            self.apply_style(worksheet, f'G{row_count}', 'attention')
            row_count += 1

        for mismatch in validation_result.get('mismatches_in_order', []):
            self.set_cell_value(worksheet, row_count, 5, mismatch['raw_name'])
            self.set_cell_value(worksheet, row_count, 7, 'no match')
            self.apply_style(worksheet, f'G{row_count}', 'attention')
            row_count += 1

        # Add txt file status
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

        # Hide rows for completed orders
        if is_completed:
            for i in range(order_row + 1, row_count):
                worksheet.row_dimensions[i].hidden = True

        return row_count

    # Complex update operations
    def copy_data_with_formatting(self, source_worksheet, start_row=2):
        """Copy data with formatting from source worksheet"""
        copied_rows = []
        
        for i, row in enumerate(source_worksheet.iter_rows(min_row=start_row, values_only=False), start=start_row):
            row_data = []
            hidden = source_worksheet.row_dimensions[i].hidden
            
            for cell in row:
                cell_fill = cell.fill
                cell_data = {
                    'value': cell.value,
                    'fill': {
                        'fill_type': cell_fill.fill_type,
                        'start_color': cell_fill.start_color,
                        'end_color': cell_fill.end_color
                    }
                }
                row_data.append(cell_data)
            
            copied_rows.append([row_data, hidden])
        
        return copied_rows

    def paste_data_with_formatting(self, worksheet, data_rows, start_row=2):
        """Paste data with formatting to worksheet"""
        for i, (row_data, hidden) in enumerate(data_rows, start=start_row):
            worksheet.row_dimensions[i].hidden = hidden
            
            for col_num, cell_data in enumerate(row_data, start=1):
                cell = worksheet.cell(row=i, column=col_num, value=cell_data['value'])
                cell.fill = PatternFill(
                    start_color=cell_data['fill']['start_color'],
                    end_color=cell_data['fill']['end_color'],
                    fill_type=cell_data['fill']['fill_type']
                )

    def update_existing_summary(self, existing_workbook, new_data_workbook, save_path):
        """Update existing summary with new data"""
        existing_sheet = existing_workbook.active
        new_sheet = new_data_workbook.active
        
        # Get new data with formatting
        new_data_rows = self.copy_data_with_formatting(new_sheet)
        num_new_rows = len(new_data_rows)
        
        if num_new_rows == 0:
            return True
            
        # Insert rows and handle data transfer
        self.insert_rows_at_top(existing_sheet, num_new_rows)
        self.paste_data_with_formatting(existing_sheet, new_data_rows, start_row=2)
        
        # Copy column widths
        for col_letter, col_dimension in new_sheet.column_dimensions.items():
            existing_sheet.column_dimensions[col_letter].width = col_dimension.width
        
        # Restore hidden states with offset
        self.restore_hidden_row_states(existing_sheet, offset=num_new_rows)
        
        return self.save_with_error_handling(existing_workbook, save_path)

    def finalize_workbook(self, worksheet, add_break_at_end=False):
        """Finalize workbook with column adjustments and optional break row"""
        self.auto_adjust_columns(worksheet)
        if add_break_at_end:
            self.add_break_row(worksheet, worksheet.max_row + 1)


if __name__ == "__main__":
    # Simplified test
    class TestConfig:
        TEXT_FILES = ['.raw.qual.txt', '.raw.seq.txt', '.seq.info.txt', '.seq.qual.txt', '.seq.txt']

    config = TestConfig()
    excel_dao = ExcelDAO(config)
    
    wb = excel_dao.create_workbook()
    ws = wb.active
    excel_dao.set_validation_headers(ws)
    
    print("Simplified ExcelDAO test completed!")