import os
import sys
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))


def inspect_reinject_list(i_numbers=None, verbose=True):
    """
    Detailed inspection of reinject list contents

    Args:
        i_numbers: List of specific I-numbers to check (default: [22055, 22056])
        verbose: Whether to show detailed debug information
    """
    from mseqauto.config import MseqConfig
    from mseqauto.core import FileSystemDAO, FolderProcessor

    # Initialize components
    config = MseqConfig()
    file_dao = FileSystemDAO(config)
    processor = FolderProcessor(file_dao, None, config)

    # Use provided I-numbers or defaults
    if not i_numbers:
        i_numbers = ['22082']

    print(f"\n=== Reinject List Inspection for I-numbers: {i_numbers} ===")

    # Get today's reinject file path
    today = datetime.now().strftime('%m-%d-%Y')
    reinject_path = f"P:\\Data\\Reinjects\\Reinject List_{today}.xlsx"

    # Check paths that will be searched
    paths_to_check = [
        'G:\\Lab\\Spreadsheets\\Individual Uploaded to ABI',
        'G:\\Lab\\Spreadsheets'
    ]

    print("\nSearching for reinject information in:")
    for path in paths_to_check:
        print(f"- {path}")
        if os.path.exists(path):
            files = [f for f in os.listdir(path) if 'reinject' in f.lower() and f.endswith('.txt')]
            if verbose:
                print(f"  Found {len(files)} reinject-related files:")
                for f in files:
                    print(f"  - {f}")

    print(f"\nChecking for Excel reinject list: {reinject_path}")
    if os.path.exists(reinject_path):
        print("Found today's reinject Excel file")
    else:
        print("Today's reinject Excel file not found")
        # Look for most recent reinject file
        reinject_dir = "P:\\Data\\Reinjects"
        if os.path.exists(reinject_dir):
            reinject_files = [f for f in os.listdir(reinject_dir) if f.startswith("Reinject List_")]
            if reinject_files:
                most_recent = sorted(reinject_files)[-1]
                reinject_path = os.path.join(reinject_dir, most_recent)
                print(f"Using most recent reinject file instead: {most_recent}")

    # Get the reinject list with detailed logging
    class DebugLogger:
        def info(self, msg):
            if verbose:
                print(f"DEBUG: {msg}")

    debug_logger = DebugLogger()
    processor.logger = debug_logger

    print("\nGetting reinject list...")
    reinject_list = processor.get_reinject_list(i_numbers, reinject_path)
    raw_reinject_list = processor.raw_reinject_list

    print(f"\nFound {len(reinject_list)} total reinjections")

    # Display the complete lists
    print("\n=== Raw Reinject List ===")
    for i, (raw, normalized) in enumerate(zip(raw_reinject_list, reinject_list)):
        print(f"\n{i+1}. RAW: {raw}")
        print(f"   NORM: {normalized}")

    # Interactive testing function
    def test_filename(filename):
        print(f"\n=== Testing filename: {filename} ===")

        # Normalize the test filename
        normalized_test = file_dao.standardize_filename_for_matching(filename)
        print(f"Normalized to: {normalized_test}")

        # Check if it's in the reinject list
        if normalized_test in reinject_list:
            idx = reinject_list.index(normalized_test)
            print(f"MATCH FOUND - Entry #{idx + 1}:")
            print(f"  Raw entry: {raw_reinject_list[idx]}")
            print(f"  Normalized entry: {reinject_list[idx]}")
        else:
            print("Not found in reinject list")

        # Check for preemptive pattern
        if processor.is_preemptive(filename):
            print("File has preemptive pattern (double well location)")

        # Check if file would go to Alternate Injections
        debug_info = processor.debug_reinject_detection(filename)
        print("\nReinject Detection Analysis:")
        for key, value in debug_info.items():
            print(f"  {key}: {value}")

    # Interactive testing loop
    print("\n=== Interactive Testing ===")
    print("Enter filenames to test why they might be flagged as reinjections")
    print("Enter 'q' to quit")

    while True:
        try:
            test_input = input("\nEnter filename to test (or 'q' to quit): ").strip()
            if test_input.lower() in ('q', 'quit', 'exit'):
                break
            if test_input:
                test_filename(test_input)
        except KeyboardInterrupt:
            break
        except Exception as e:
            print(f"Error testing file: {e}")

    return reinject_list, raw_reinject_list

if __name__ == "__main__":
    # Allow command-line specification of I-numbers
    i_numbers = sys.argv[1:] if len(sys.argv) > 1 else None
    inspect_reinject_list(i_numbers)