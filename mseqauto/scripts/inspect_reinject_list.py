import os
import sys
import tkinter as tk
from tkinter import filedialog
import re
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))


def inspect_reinject_list(folder_path=None, i_numbers=None, show_normalized=True):
    """
    Helper function to display and inspect the reinject list
    
    Args:
        folder_path (str, optional): Path to data folder to find I-numbers from
        i_numbers (list, optional): Specific I-numbers to check for reinjections
        show_normalized (bool): Whether to show normalized versions alongside originals
        
    Returns:
        None (prints information to console)
    """
    from datetime import datetime
    import os
    import re
    
    # Import required modules from MseqAuto
    from mseqauto.config import MseqConfig
    from mseqauto.core import FileSystemDAO, FolderProcessor
    
    # Initialize components
    config = MseqConfig()
    file_dao = FileSystemDAO(config)
    processor = FolderProcessor(file_dao, None, config)
    
    # Create output file with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"reinject_inspection_{timestamp}.txt"
    
    # Function to write to both console and file
    def log(message):
        print(message)
        with open(output_file, "a") as f:
            f.write(message + "\n")
    
    log(f"\n=== Reinject List Inspection ===")
    log(f"Output saved to: {output_file}")
    
    # Get I-numbers from folder if provided
    if folder_path and not i_numbers:
        i_numbers, _ = file_dao.get_folders_with_inumbers(folder_path)
        log(f"Found I-numbers from folder: {i_numbers}")
    
    # If no I-numbers provided or found, use some defaults
    if not i_numbers:
        # Use recent I-numbers based on most recent from Data/Individuals 
        most_recent = file_dao.get_most_recent_inumber('P:\\Data\\Individuals')
        if most_recent:
            # Create a range of I-numbers around the most recent
            most_recent_int = int(most_recent)
            i_numbers = [str(most_recent_int - i) for i in range(5, -5, -1)]
            log(f"Using range of I-numbers around most recent ({most_recent}): {i_numbers}")
        else:
            # Fallback to hardcoded numbers if we can't find most recent
            i_numbers = [str(21000 + i) for i in range(10)]
            log(f"Using default I-numbers: {i_numbers}")
    
    # Get reinject file path
    reinject_path = f"P:\\Data\\Reinjects\\Reinject List_{datetime.now().strftime('%m-%d-%Y')}.xlsx"
    if os.path.exists(reinject_path):
        log(f"Using reinject file: {reinject_path}")
    else:
        log(f"Warning: Reinject file not found at {reinject_path}")
        # Try to find alternative reinject file
        reinject_dir = "P:\\Data\\Reinjects"
        if os.path.exists(reinject_dir):
            reinject_files = [f for f in os.listdir(reinject_dir) if f.startswith("Reinject List_")]
            if reinject_files:
                # Use the most recent reinject file
                reinject_path = os.path.join(reinject_dir, sorted(reinject_files)[-1])
                log(f"Using alternative reinject file: {reinject_path}")
    
    # Get the reinject list
    log("\nLoading reinject list...")
    reinject_list = processor.get_reinject_list(i_numbers, reinject_path)
    raw_reinject_list = processor.raw_reinject_list
    
    # Display statistics
    log(f"Found {len(reinject_list)} reinjections")
    
    # Check for pre-emptive reinject markers
    preemptive_count = sum(1 for item in raw_reinject_list if "{!P}" in item)
    log(f"Pre-emptive reinjections: {preemptive_count}")
    
    # Display full lists
    log("\n=== Raw Reinject List ===")
    for i, item in enumerate(raw_reinject_list):
        if show_normalized:
            normalized = file_dao.standardize_filename_for_matching(item)
            log(f"{i+1:3}. RAW: {item}")
            log(f"     NORM: {normalized}")
        else:
            log(f"{i+1:3}. {item}")
    
    # Check for common patterns in reinject list
    log("\n=== Pattern Analysis ===")
    
    # Check for double well location pattern {XXX}{YYY}
    double_well_pattern = re.compile(r'^{\d+[A-Z]}{\d+[A-Z]}')
    double_well_count = sum(1 for item in raw_reinject_list if double_well_pattern.match(item))
    log(f"Items with double well location pattern: {double_well_count}")
    
    # Compile regex patterns we'll use for testing
    patterns = {
        'double_well': double_well_pattern,
        'pcr_number': re.compile(r'{pcr(\d+).+}', re.IGNORECASE),
        'preemptive': re.compile(r'{!P}'),
        'well_location': re.compile(r'{\d+[A-Z]}')
    }
    
    # Function to test if a filename would be considered a reinject
    def test_file(filename):
        log(f"\nTesting file: {filename}")
        
        # Check different methods that might flag a file as reinject
        
        # 1. Check against reinject list
        normalized_name = file_dao.standardize_filename_for_matching(filename)
        normalized_reinjects = [file_dao.standardize_filename_for_matching(r) for r in reinject_list]
        in_reinject_list = normalized_name in normalized_reinjects
        log(f"1. In reinject list: {in_reinject_list}")
        
        # 2. Check for double well location pattern
        has_double_well = bool(patterns['double_well'].match(filename))
        log(f"2. Has double well pattern: {has_double_well}")
        
        # 3. Check for pre-emptive marker
        is_preemptive = bool(patterns['preemptive'].search(filename))
        log(f"3. Is marked as pre-emptive: {is_preemptive}")
        
        # 4. Check for PCR number
        pcr_match = patterns['pcr_number'].search(filename)
        has_pcr = bool(pcr_match)
        pcr_number = pcr_match.group(1) if pcr_match else None
        log(f"4. Has PCR number: {has_pcr} {f'(PCR{pcr_number})' if pcr_number else ''}")
        
        # 5. Check for regular well location
        has_well = bool(patterns['well_location'].search(filename))
        log(f"5. Has well location: {has_well}")
        
        # 6. Check for suffix patterns
        has_rtx_suffix = "_RTI" in filename
        has_premixed_suffix = "_Premixed" in filename
        log(f"6. Has special suffix: RTI={has_rtx_suffix}, Premixed={has_premixed_suffix}")
        
        # Overall determination
        is_reinject = in_reinject_list or has_double_well
        log(f"\nWould be considered a reinject: {is_reinject}")
        
        # How it would be processed
        if is_reinject:
            if is_preemptive:
                log(f"Would be directed to: main folder (pre-emptive reinject)")
            else:
                log(f"Would be directed to: Alternate Injections folder (regular reinject)")
        else:
            log(f"Would be directed to: main folder (not a reinject)")
        
        # Normalized filename
        log(f"\nOriginal filename: {filename}")
        log(f"Normalized filename: {normalized_name}")
        
        # If in reinject list, show matching entry
        if in_reinject_list:
            match_indices = [i for i, norm in enumerate(normalized_reinjects) if norm == normalized_name]
            for idx in match_indices:
                log(f"\nMatched with reinject #{idx+1}: {raw_reinject_list[idx]}")
                log(f"  Normalized to: {normalized_reinjects[idx]}")
        
        # Close match search - to catch near matches
        if not in_reinject_list:
            log("\nSearching for close matches in reinject list...")
            
            # Check for partial matches
            close_matches = []
            for i, norm in enumerate(normalized_reinjects):
                # Check if normalized name is a substring of or contains a reinject entry
                if normalized_name in norm or norm in normalized_name:
                    close_matches.append((i, norm, raw_reinject_list[i]))
            
            if close_matches:
                log(f"Found {len(close_matches)} close matches:")
                for i, norm, raw in close_matches:
                    log(f"  #{i+1}: {raw}")
                    log(f"     Normalized to: {norm}")
            else:
                log("No close matches found")
    
    # Check if command-line argument was provided
    import sys
    if len(sys.argv) > 1 and sys.argv[1] != folder_path:
        test_filename = sys.argv[1]
        test_file(test_filename)
    
    # Interactive testing loop
    log("\n=== Interactive File Testing ===")
    log("Enter filenames to test if they would be considered a reinject")
    log("Type 'quit', 'exit', or 'q' to end testing")
    
    while True:
        try:
            test_input = input("\nEnter filename to test (or 'quit' to exit): ")
            if test_input.lower() in ('quit', 'exit', 'q'):
                log("\nEnding interactive testing")
                break
            
            if test_input:
                test_file(test_input)
            else:
                log("No filename entered, try again or type 'quit' to exit")
        except KeyboardInterrupt:
            log("\nTesting interrupted by user")
            break
        except Exception as e:
            log(f"Error testing file: {e}")
    
    log(f"\nInspection complete. Results saved to {output_file}")
    print(f"\nInspection complete. Results saved to {output_file}")
    return reinject_list, raw_reinject_list  # Return these for potential further analysis

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        # If folder path is provided
        inspect_reinject_list(folder_path=sys.argv[1])
    else:
        # Interactive mode
        folder_path = input("Enter folder path to inspect (or press Enter to use defaults): ")
        if folder_path:
            inspect_reinject_list(folder_path=folder_path)
        else:
            inspect_reinject_list()