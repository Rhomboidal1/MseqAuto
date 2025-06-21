# end_of_day_organize.py
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import re
from datetime import datetime

# Add parent directory to PYTHONPATH for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

def get_folder_from_user():
    """Get today's data folder from user"""
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    root.update()
    
    folder_path = filedialog.askdirectory(
        title="Select today's data folder to organize",
        mustexist=True
    )
    
    root.destroy()
    return folder_path

class EndOfDayOrganizer:
    def __init__(self, data_folder, dry_run=True):
        # Import after path setup
        from mseqauto.config import MseqConfig #type: ignore
        from mseqauto.core import FileSystemDAO #type: ignore
        from mseqauto.utils import setup_logger #type: ignore
        
        self.data_folder = data_folder
        self.dry_run = dry_run
        self.config = MseqConfig()
        
        # Setup logger
        script_dir = os.path.dirname(os.path.abspath(__file__))
        log_dir = os.path.join(script_dir, "logs")
        self.logger = setup_logger("end_of_day_organize", log_dir=log_dir)
        
        # Initialize file DAO
        self.file_dao = FileSystemDAO(self.config, logger=self.logger)
        
        # Define target directories - correct path construction with normalized paths
        data_parent = os.path.dirname(data_folder)  # Gets P:/Data  
        p_drive = os.path.dirname(data_parent)      # Gets P:
        
        self.individuals_dir = os.path.normpath(os.path.join(data_parent, "Individuals"))  # P:/Data/Individuals
        self.plates_dir = os.path.normpath(os.path.join(data_parent, "Plates"))            # P:/Data/Plates  
        self.pcr_base_dir = os.path.normpath(os.path.join(p_drive, "PCR"))                 # P:/PCR
        
        self.logger.info(f"Initializing End of Day Organizer")
        self.logger.info(f"Source folder: {data_folder}")
        self.logger.info(f"Individuals target: {self.individuals_dir}")
        self.logger.info(f"Plates target: {self.plates_dir}")
        self.logger.info(f"PCR base: {self.pcr_base_dir}")
        self.logger.info(f"Dry run mode: {dry_run}")

    def find_pcr_parent_folder(self, pcr_folder_name):
        """Find the matching PCR parent folder in PCR/"""
        # Handle both formats: FB-PCR3046 and FB-PCR3046_149703
        # Extract PCR number using a more flexible pattern
        pcr_match = re.search(r'fb-pcr(\d+)', pcr_folder_name.lower())
        if not pcr_match:
            self.logger.warning(f"Could not extract PCR number from: {pcr_folder_name}")
            return None
            
        pcr_number = pcr_match.group(1)
        expected_parent = f"FB-PCR{pcr_number}"
        parent_path = os.path.join(self.pcr_base_dir, expected_parent)
        
        if os.path.exists(parent_path):
            self.logger.info(f"Found PCR parent folder: {parent_path}")
            return parent_path
        else:
            self.logger.warning(f"PCR parent folder not found: {parent_path}")
            return None

    def get_folders_to_move(self):
        """Identify all folders that need to be moved"""
        moves = {
            'individuals': [],
            'plates': [], 
            'pcr': []
        }
        
        if not os.path.exists(self.data_folder):
            self.logger.error(f"Data folder does not exist: {self.data_folder}")
            return moves
            
        for item in os.listdir(self.data_folder):
            item_path = os.path.join(self.data_folder, item)
            
            # Skip files and special folders
            if not os.path.isdir(item_path):
                continue
            if item.lower() in ['ind not ready', 'plate not ready', 'zip dump']:
                continue
            if item.endswith('.xlsx') or item.endswith('.txt'):
                continue
                
            # Check folder type
            if self.config.REGEX_PATTERNS['bioi_folder'].search(item.lower()):
                moves['individuals'].append((item_path, item))
                self.logger.info(f"Found BioI folder: {item}")
                
            elif re.match(r'^p\d+.+', item.lower()):
                moves['plates'].append((item_path, item))
                self.logger.info(f"Found Plate folder: {item}")
                
            elif re.match(r'^fb-pcr\d+', item.lower()):  # More flexible PCR matching
                parent_folder = self.find_pcr_parent_folder(item)
                if parent_folder:
                    moves['pcr'].append((item_path, item, parent_folder))
                    self.logger.info(f"Found PCR folder: {item} -> {parent_folder}")
                else:
                    self.logger.warning(f"Skipping PCR folder (no parent): {item}")
            else:
                self.logger.info(f"Unrecognized folder type, skipping: {item}")
                
        return moves

    def safe_move_folder(self, source_path, destination_parent, folder_name):
        """Safely move a folder, handling merging if destination exists"""
        destination_path = os.path.join(destination_parent, folder_name)
        
        # Normalize paths for consistent display
        source_display = os.path.normpath(source_path).replace('\\', '/')
        dest_display = os.path.normpath(destination_path).replace('\\', '/')
        
        if self.dry_run:
            print(f"  Would move: {source_display}")
            print(f"         to: {dest_display}")
            self.logger.info(f"DRY RUN: Would move {source_path} -> {destination_path}")
            if not os.path.exists(destination_parent):
                print(f"  WARNING: Destination parent does not exist: {os.path.normpath(destination_parent).replace('\\', '/')}")
                self.logger.warning(f"DRY RUN: Destination parent does not exist: {destination_parent}")
                return False
            return True
            
        try:
            # Check if destination parent exists
            if not os.path.exists(destination_parent):
                self.logger.error(f"Destination parent does not exist: {destination_parent}")
                return False
                
            # Use the file_dao move_folder method which handles merging
            print(f"  Moving: {folder_name}")
            success = self.file_dao.move_folder(source_path, destination_path)
            
            if success:
                print(f"  ✓ Successfully moved: {folder_name}")
                self.logger.info(f"Successfully moved: {folder_name}")
                return True
            else:
                print(f"  ✗ Failed to move: {folder_name}")
                self.logger.error(f"Failed to move: {folder_name}")
                return False
                
        except Exception as e:
            self.logger.error(f"Error moving {folder_name}: {e}")
            return False

    def verify_destinations_exist(self):
        """Verify that all destination directories exist"""
        destinations = [
            ("Individuals", self.individuals_dir),
            ("Plates", self.plates_dir),
            ("PCR", self.pcr_base_dir)
        ]
        
        missing = []
        for name, path in destinations:
            if not os.path.exists(path):
                missing.append((name, path))
                self.logger.error(f"Missing destination directory {name}: {path}")
            else:
                self.logger.info(f"Verified destination {name}: {path}")
        
        return missing

    def organize_folders(self):
        """Main method to organize all folders"""
        print("Starting end-of-day folder organization...")
        self.logger.info("Starting end-of-day folder organization")
        
        # First verify destinations exist
        missing_destinations = self.verify_destinations_exist()
        if missing_destinations and not self.dry_run:
            print("ERROR: Cannot proceed - missing destination directories:")
            self.logger.error("Cannot proceed - missing destination directories:")
            for name, path in missing_destinations:
                print(f"  {name}: {os.path.normpath(path).replace('\\', '/')}")
                self.logger.error(f"  {name}: {path}")
            return False
        
        # Get all folders to move
        moves = self.get_folders_to_move()
        
        total_folders = len(moves['individuals']) + len(moves['plates']) + len(moves['pcr'])
        
        if total_folders == 0:
            print("No folders found to organize")
            self.logger.info("No folders found to organize")
            return True
            
        print(f"\nFound {total_folders} folders to organize:")
        print(f"  - {len(moves['individuals'])} BioI folders -> Individuals")
        print(f"  - {len(moves['plates'])} Plate folders -> Plates") 
        print(f"  - {len(moves['pcr'])} PCR folders -> PCR")
        
        self.logger.info(f"Found {total_folders} folders to organize:")
        self.logger.info(f"  - {len(moves['individuals'])} BioI folders -> Individuals")
        self.logger.info(f"  - {len(moves['plates'])} Plate folders -> Plates") 
        self.logger.info(f"  - {len(moves['pcr'])} PCR folders -> PCR")
        
        if self.dry_run:
            print("\n--- DRY RUN MODE ---")
        
        # Track results
        results = {'success': 0, 'failed': 0}
        
        # Move BioI folders to Individuals
        if moves['individuals']:
            print(f"\nMoving BioI folders to Individuals:")
        for source_path, folder_name in moves['individuals']:
            if self.safe_move_folder(source_path, self.individuals_dir, folder_name):
                results['success'] += 1
            else:
                results['failed'] += 1
                
        # Move Plate folders to Plates  
        if moves['plates']:
            print(f"\nMoving Plate folders to Plates:")
        for source_path, folder_name in moves['plates']:
            if self.safe_move_folder(source_path, self.plates_dir, folder_name):
                results['success'] += 1
            else:
                results['failed'] += 1
                
        # Move PCR folders to their parent folders
        if moves['pcr']:
            print(f"\nMoving PCR folders to PCR:")
        for source_path, folder_name, parent_folder in moves['pcr']:
            if self.safe_move_folder(source_path, parent_folder, folder_name):
                results['success'] += 1
            else:
                results['failed'] += 1
        
        # Summary
        print(f"\nOrganization complete: {results['success']} successful, {results['failed']} failed")
        self.logger.info(f"Organization complete: {results['success']} successful, {results['failed']} failed")
        
        return results['failed'] == 0

def main():
    """Main function"""
    # Get today's data folder
    data_folder = get_folder_from_user()
    
    if not data_folder:
        print("No folder selected, exiting")
        return
        
    print(f"Selected folder: {data_folder}")
    
    # Ask about dry run
    root = tk.Tk()
    root.withdraw()
    
    dry_run = messagebox.askyesno(
        "Dry Run Mode", 
        "Run in dry-run mode first?\n\n"
        "Dry run will show what would be moved without actually moving anything.\n"
        "Recommended for first time use."
    )
    
    root.destroy()
    
    # Initialize organizer
    organizer = EndOfDayOrganizer(data_folder, dry_run=dry_run)
    
    # Run organization
    success = organizer.organize_folders()
    
    if dry_run:
        # Ask if they want to run for real
        root = tk.Tk()
        root.withdraw()
        
        if messagebox.askyesno("Run for Real?", "Dry run complete. Run the actual organization now?"):
            organizer.dry_run = False
            organizer.logger.info("Running actual organization after dry run")
            success = organizer.organize_folders()
        
        root.destroy()
    
    # Final message
    if success:
        print("Organization completed successfully!")
    else:
        print("Organization completed with some errors. Check the log for details.")

if __name__ == "__main__":
    main()