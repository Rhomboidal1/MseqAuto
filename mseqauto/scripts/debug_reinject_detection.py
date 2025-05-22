import os
import sys
import tkinter as tk
from tkinter import filedialog
import re
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

def debug_reinject_detection(folder_path=None, i_numbers=None, target_files=None):
    """
    Advanced debugging function to trace exactly why files are being marked as reinjections
    
    Args:
        folder_path (str, optional): Path to data folder
        i_numbers (list, optional): Specific I-numbers to check
        target_files (list, optional): Specific filenames to trace through detection logic
        
    Returns:
        None (prints detailed debugging information)
    """
    from datetime import datetime
    import os
    import re
    import inspect
    
    # Import required modules from MseqAuto
    from mseqauto.config import MseqConfig
    from mseqauto.core import FileSystemDAO, FolderProcessor
    
    # Create output file with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"reinject_debug_{timestamp}.txt"
    
    # Function to write to both console and file
    def log(message):
        print(message)
        with open(output_file, "a") as f:
            f.write(message + "\n")
    
    log(f"\n=== REINJECT DETECTION DEBUGGING ===")
    log(f"Output saved to: {output_file}")
    
    # Initialize components
    config = MseqConfig()
    file_dao = FileSystemDAO(config)
    
    # Create a special processor with debug hooks
    class DebugProcessor(FolderProcessor):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.debug_log = log
        
        def sort_customer_file(self, file_path, order_key):
            """Override to trace how files are processed"""
            file_name = os.path.basename(file_path)
            self.debug_log(f"\n--- TRACING SORT PATH FOR: {file_name} ---")
            
            # Now trace the exact path through the original method
            result = super().sort_customer_file(file_path, order_key)
            return result
        
        def _move_file_to_destination(self, file_path, destination_folder, normalized_name):
            """Override to trace reinject detection logic"""
            file_name = os.path.basename(file_path)
            
            # Check if this is a target file we want to debug
            if target_files and not any(tf in file_name for tf in target_files):
                return super()._move_file_to_destination(file_path, destination_folder, normalized_name)
            
            self.debug_log(f"\n--- REINJECT DETECTION FOR: {file_name} ---")
            self.debug_log(f"Normalized name: {normalized_name}")
            
            # Check if file is a reinject by presence in reinject list
            is_reinject = False
            raw_name = None
            
            # Check in reinject list
            if hasattr(self, 'reinject_list') and self.reinject_list:
                # Get full normalized list for comparison
                normalized_reinject_list = [
                    self.file_dao.standardize_filename_for_matching(r) for r in self.reinject_list
                ]
                
                # First check for exact matches
                if normalized_name in normalized_reinject_list:
                    is_reinject = True
                    idx = normalized_reinject_list.index(normalized_name)
                    raw_name = self.raw_reinject_list[idx] if hasattr(self, 'raw_reinject_list') else self.reinject_list[idx]
                    self.debug_log(f"MATCH FOUND in reinject list at index {idx}")
                    self.debug_log(f"Raw reinject list entry: {raw_name}")
                    self.debug_log(f"Normalized reinject entry: {normalized_reinject_list[idx]}")
                else:
                    self.debug_log(f"Not found in reinject list (exact match)")
                    
                    # Check for partial matches (helpful for debugging)
                    close_matches = []
                    for i, norm in enumerate(normalized_reinject_list):
                        if normalized_name in norm or norm in normalized_name:
                            close_matches.append((i, norm, self.reinject_list[i]))
                    
                    if close_matches:
                        self.debug_log(f"Found {len(close_matches)} close matches in reinject list:")
                        for i, norm, raw in close_matches:
                            self.debug_log(f"  #{i+1}: {raw}")
                            self.debug_log(f"     Normalized to: {norm}")
            
            # Check for double well pattern (triggers reinject logic)
            double_well_pattern = re.compile(r'^{\d+[A-Z]}{(\d+[A-Z])}')
            
            has_double_well = bool(double_well_pattern.match(file_name))
            if has_double_well:
                is_reinject = True
                self.debug_log(f"MATCH FOUND for double well pattern: {file_name}")
            
            # Check processor's reinject detection implementation
            # Dump the actual code being used for clarity
            self.debug_log(f"\nActual _move_file_to_destination code being executed:")
            source_code = inspect.getsource(FolderProcessor._move_file_to_destination)
            self.debug_log(source_code)
            
            self.debug_log(f"Final reinject determination: {is_reinject}")
            
            # Continue with original method
            return super()._move_file_to_destination(file_path, destination_folder, normalized_name)
    
    # Use our debug processor
    processor = DebugProcessor(file_dao, None, config, logger=log)
    
    # Get I-numbers from folder if provided
    if folder_path and not i_numbers:
        i_numbers, _ = file_dao.get_folders_with_inumbers(folder_path)
        log(f"Found I-numbers from folder: {i_numbers}")
    
    # If no I-numbers provided or found, use some defaults
    if not i_numbers:
        # Use recent I-numbers
        most_recent = file_dao.get_most_recent_inumber('P:\\Data\\Individuals')
        if most_recent:
            most_recent_int = int(most_recent)
            i_numbers = [str(most_recent_int - i) for i in range(5, -5, -1)]
            log(f"Using range of I-numbers around most recent ({most_recent}): {i_numbers}")
        else:
            i_numbers = [str(21000 + i) for i in range(10)]
            log(f"Using default I-numbers: {i_numbers}")
    
    # FIXED: Get reinject file path with multiple date format attempts
    # First try with today's date in multiple formats
    today = datetime.now()
    month = today.month  # Get as integer
    day = today.day      # Get as integer
    year = today.year    # Get as integer
    
    # Try multiple date formats and find the most recent file if needed
    reinject_path = None
    reinject_dir = "P:\\Data\\Reinjects"
    
    # Function to find the most recent reinject file
    def find_most_recent_reinject_file():
        if not os.path.exists(reinject_dir):
            return None
        
        reinject_files = [f for f in os.listdir(reinject_dir) if f.startswith("Reinject List_")]
        if not reinject_files:
            return None
            
        # Sort by modification time (newest first)
        sorted_files = sorted(
            [(os.path.join(reinject_dir, f), os.path.getmtime(os.path.join(reinject_dir, f))) 
             for f in reinject_files],
            key=lambda x: x[1],
            reverse=True
        )
        
        return sorted_files[0][0] if sorted_files else None
    
    # Try with different date formats (use string formatting instead of strftime)
    date_formats = [
        f"{month:02d}-{day:02d}-{year}",   # Format: MM-DD-YYYY (e.g., 04-21-2025)
        f"{month}-{day}-{year}",           # Format: M-D-YYYY (e.g., 4-21-2025)
        f"{month:02d}-{day:02d}-{year % 100:02d}",  # Format: MM-DD-YY (e.g., 04-21-25)
        f"{month}-{day}-{year % 100}"      # Format: M-D-YY (e.g., 4-21-25)
    ]
    
    log("\nChecking for reinject files with formats:")
    for fmt in date_formats:
        log(f"- Reinject List_{fmt}.xlsx")
    
    # Try all date formats first
    for date_format in date_formats:
        test_path = os.path.join(reinject_dir, f"Reinject List_{date_format}.xlsx")
        if os.path.exists(test_path):
            reinject_path = test_path
            log(f"Found reinject file with format: {test_path}")
            break
    
    # If not found, use most recent file
    if not reinject_path:
        reinject_path = find_most_recent_reinject_file()
        if reinject_path:
            log(f"Using most recent reinject file: {reinject_path}")
        else:
            log(f"Warning: No reinject files found in {reinject_dir}")
    
    # Get the reinject list
    log("\nLoading reinject list...")
    
    if reinject_path:
        reinject_list = processor.get_reinject_list(i_numbers, reinject_path)
        raw_reinject_list = processor.raw_reinject_list
        log(f"Loaded {len(reinject_list)} items in reinject list")
    else:
        log("No reinject file found. Continuing with empty reinject list.")
        reinject_list = []
        raw_reinject_list = []
    
    # Now test specific files against our mock processor
    test_files = target_files or []
    if not test_files:
        log("\n=== Interactive File Testing ===")
        log("Enter filenames to test against reinject detection logic")
        log("Type 'quit', 'exit', or 'q' to end testing")
        
        while True:
            try:
                test_input = input("\nEnter filename to test (or 'quit' to exit): ")
                if test_input.lower() in ('quit', 'exit', 'q'):
                    log("\nEnding interactive testing")
                    break
                
                if test_input:
                    test_files.append(test_input)
                    # Deep trace through the actual file processing logic
                    mock_path = os.path.join(folder_path or "P:\\Data\\Testing\\MOCK", test_input)
                    mock_destination = "P:\\Data\\Testing\\MOCK\\DESTINATION"
                    log(f"Simulating processing {test_input}...")
                    processor._move_file_to_destination(mock_path, mock_destination, 
                                                       file_dao.standardize_filename_for_matching(test_input))
                else:
                    log("No filename entered, try again or type 'quit' to exit")
            except KeyboardInterrupt:
                log("\nTesting interrupted by user")
                break
            except Exception as e:
                log(f"Error testing file: {e}")
                import traceback
                log(traceback.format_exc())
    
    # Extra step: check for memory or stateful tracking issues
    log("\n=== Testing for State Persistence Problems ===")
    
    # Test processing the non-reinject file, then the reinject file, then the non-reinject again
    # This will reveal if there's a state persistence issue
    test_non_reinject = "{09A}1_FWDPARP1_Premixed.ab1"
    test_reinject = "{01A}{09A}1_FWDPARP1_Premixed{2_18}{I-22082}.ab1"
    
    log("\nTesting sequence to detect state persistence issues:")
    
    log(f"\n1. First testing non-reinject file: {test_non_reinject}")
    mock_path = os.path.join(folder_path or "P:\\Data\\Testing\\MOCK", test_non_reinject)
    mock_destination = "P:\\Data\\Testing\\MOCK\\DESTINATION"
    norm_name = file_dao.standardize_filename_for_matching(test_non_reinject)
    processor._move_file_to_destination(mock_path, mock_destination, norm_name)
    
    log(f"\n2. Now testing reinject file: {test_reinject}")
    mock_path = os.path.join(folder_path or "P:\\Data\\Testing\\MOCK", test_reinject)
    processor._move_file_to_destination(mock_path, mock_destination, 
                                       file_dao.standardize_filename_for_matching(test_reinject))
    
    log(f"\n3. Testing non-reinject file again: {test_non_reinject}")
    mock_path = os.path.join(folder_path or "P:\\Data\\Testing\\MOCK", test_non_reinject)
    processor._move_file_to_destination(mock_path, mock_destination, norm_name)
    
    # Finally, check for duplicates or conflicts in normalized names
    log("\n=== Checking reinject list for normalization conflicts ===")
    norm_to_raw = {}
    for i, item in enumerate(raw_reinject_list):
        norm = file_dao.standardize_filename_for_matching(item)
        if norm in norm_to_raw:
            log(f"CONFLICT FOUND: '{norm}'")
            log(f"  Entry #{norm_to_raw[norm][0]+1}: {norm_to_raw[norm][1]}")
            log(f"  Entry #{i+1}: {item}")
        else:
            norm_to_raw[norm] = (i, item)
            
    log(f"\nDebugging complete. Results saved to {output_file}")
    print(f"\nDebugging complete. Results saved to {output_file}")
if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        # If folder path is provided
        debug_reinject_detection(folder_path=sys.argv[1])
    else:
        # Interactive mode
        folder_path = input("Enter folder path to inspect (or press Enter to use defaults): ")
        target_files = input("Enter specific filenames to debug (comma-separated, or press Enter for none): ")
        
        if target_files:
            target_files = [f.strip() for f in target_files.split(',')]
        else:
            target_files = None
            
        debug_reinject_detection(folder_path=folder_path if folder_path else None, 
                              target_files=target_files)