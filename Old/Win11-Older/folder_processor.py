# folder_processor.py
import os
import re
import numpy as np

class FolderProcessor:
    def __init__(self, file_dao, ui_automation, config, logger=None):
        self.file_dao = file_dao
        self.ui_automation = ui_automation
        self.config = config
        self.logger = logger or print  # Default to print if no logger provided
    
    def is_mseq_processed(self, folder):
        """Check if a folder has been processed by mSeq"""
        current_artifacts = []
        for item in os.listdir(folder):
            if item in self.config.MSEQ_ARTIFACTS:
                current_artifacts.append(item)
        return set(current_artifacts) == self.config.MSEQ_ARTIFACTS
    
    def has_output_files(self, folder):
        """Check if folder has all 5 expected output files"""
        count = 0
        for item in os.listdir(folder):
            if os.path.isfile(os.path.join(folder, item)):
                for extension in self.config.TEXT_FILES:
                    if item.endswith(extension):
                        count += 1
        return count == 5
    
    def check_order_status(self, folder):
        """Check if an order folder has been processed, has braces, and has ab1 files"""
        was_mseqed = False
        has_braces = False
        has_ab1_files = False
        
        # Check for mSeq artifacts
        current_artifacts = []
        for item in os.listdir(folder):
            if item in self.config.MSEQ_ARTIFACTS:
                current_artifacts.append(item)
            
            # Check for braces in ab1 files
            if item.endswith('.ab1'):
                has_ab1_files = True
                if '{' in item or '}' in item:
                    has_braces = True
        
        was_mseqed = set(current_artifacts) == self.config.MSEQ_ARTIFACTS
        return was_mseqed, has_braces, has_ab1_files
    
    def get_order_number(self, folder_name):
        """Extract order number from folder name"""
        match = re.search('_\\d+', folder_name)
        if match:
            return re.search('\\d+', match.group(0)).group(0)
        return ''
    
    def get_i_number(self, folder_name):
        """Extract I number from folder name"""
        match = re.search('bioi-\\d+', str(folder_name).lower())
        if match:
            bio_string = match.group(0)
            return re.search('\\d+', bio_string).group(0)
        return None
    
    def get_destination_for_order(self, order_folder, base_path):
        """Determine the correct destination for an order folder"""
        folder_name = os.path.basename(order_folder)
        i_num = self.get_i_number(folder_name)
        
        if i_num:
            # Navigate up one level to get to the day folder
            day_data_path = os.path.dirname(base_path)
            
            # Search for the matching I number folder
            for item in os.listdir(day_data_path):
                if (re.search(f'bioi-{i_num}', item.lower()) and 
                    not re.search('reinject', item.lower())):
                    return os.path.join(day_data_path, item)
        
        # If no matching folder found, return the day data path
        return os.path.dirname(base_path)
    
    def process_bio_folder(self, folder):
        """Process a BioI folder (specialized for IND)"""
        self.logger(f"Processing BioI folder: {os.path.basename(folder)}")
        
        # Get all order folders in this BioI folder
        order_folders = self.file_dao.get_folders(folder, r'bioi-\d+_.+_\d+')
        
        for order_folder in order_folders:
            # Skip Andreev's orders for mSeq processing
            if 'andreev' in order_folder.lower():
                continue
                
            order_number = self.get_order_number(os.path.basename(order_folder))
            ab1_files = self.file_dao.get_files_by_extension(order_folder, '.ab1')
            
            # Check order status
            was_mseqed, has_braces, has_ab1_files = self.check_order_status(order_folder)
            
            if not was_mseqed and not has_braces:
                # Process if we have the right number of ab1 files
                expected_count = self._get_expected_file_count(order_number)
                
                if len(ab1_files) == expected_count:
                    if has_ab1_files:
                        self.ui_automation.process_folder(order_folder)
                        self.logger(f"mSeq completed: {os.path.basename(order_folder)}")
                else:
                    # Move to Not Ready folder if incomplete
                    not_ready_path = os.path.join(os.path.dirname(folder), "IND Not Ready")
                    self.file_dao.create_folder_if_not_exists(not_ready_path)
                    self.file_dao.move_folder(order_folder, not_ready_path)
                    self.logger(f"Order moved to Not Ready: {os.path.basename(order_folder)}")
    
    def process_plate_folder(self, folder):
        """Process a plate folder"""
        self.logger(f"Processing plate folder: {os.path.basename(folder)}")
        
        # Check if folder contains FSA files (skip if it does)
        if self.file_dao.contains_file_type(folder, '.fsa'):
            self.logger(f"Skipping folder with FSA files: {os.path.basename(folder)}")
            return
        
        # Check if already processed
        was_mseqed = self.is_mseq_processed(folder)
        
        if not was_mseqed:
            # Process with mSeq
            self.ui_automation.process_folder(folder)
            self.logger(f"mSeq completed: {os.path.basename(folder)}")
    
    def process_wildcard_folder(self, folder):
        """Process any folder"""
        self.logger(f"Processing folder: {os.path.basename(folder)}")
        
        # Skip folders with FSA files
        if self.file_dao.contains_file_type(folder, '.fsa'):
            self.logger(f"Skipping folder with FSA files: {os.path.basename(folder)}")
            return
        
        # Check if already processed
        was_mseqed = self.is_mseq_processed(folder)
        
        if not was_mseqed:
            # Process with mSeq
            self.ui_automation.process_folder(folder)
            self.logger(f"mSeq completed: {os.path.basename(folder)}")
    
    def process_order_folder(self, order_folder, data_folder_path):
        """Process an order folder"""
        self.logger(f"Processing order folder: {os.path.basename(order_folder)}")
        
        # Skip Andreev's orders for mSeq processing
        if 'andreev' in order_folder.lower():
            order_number = self.get_order_number(os.path.basename(order_folder))
            ab1_files = self.file_dao.get_files_by_extension(order_folder, '.ab1')
            expected_count = self._get_expected_file_count(order_number)
            
            # For Andreev's orders, just check if complete to move back if needed
            if len(ab1_files) == expected_count:
                # If we're processing from IND Not Ready, move it back
                if os.path.basename(os.path.dirname(order_folder)) == "IND Not Ready":
                    destination = self.get_destination_for_order(order_folder, data_folder_path)
                    self.file_dao.move_folder(order_folder, destination)
                    self.logger(f"Andreev's order moved back: {os.path.basename(order_folder)}")
            return
            
        order_number = self.get_order_number(os.path.basename(order_folder))
        ab1_files = self.file_dao.get_files_by_extension(order_folder, '.ab1')
        
        # Check order status
        was_mseqed, has_braces, has_ab1_files = self.check_order_status(order_folder)
        
        # Process based on status
        if not was_mseqed and not has_braces:
            expected_count = self._get_expected_file_count(order_number)
            
            if len(ab1_files) == expected_count and has_ab1_files:
                self.ui_automation.process_folder(order_folder)
                self.logger(f"mSeq completed: {os.path.basename(order_folder)}")
                
                # If processing from IND Not Ready, move it back
                if os.path.basename(os.path.dirname(order_folder)) == "IND Not Ready":
                    destination = self.get_destination_for_order(order_folder, data_folder_path)
                    self.file_dao.move_folder(order_folder, destination)
            else:
                # Move to Not Ready if incomplete
                not_ready_path = os.path.join(os.path.dirname(data_folder_path), "IND Not Ready")
                self.file_dao.create_folder_if_not_exists(not_ready_path)
                self.file_dao.move_folder(order_folder, not_ready_path)
                self.logger(f"Order moved to Not Ready: {os.path.basename(order_folder)}")
        
        # If already mSeqed but in IND Not Ready, move it back
        elif was_mseqed and os.path.basename(os.path.dirname(order_folder)) == "IND Not Ready":
            destination = self.get_destination_for_order(order_folder, data_folder_path)
            self.file_dao.move_folder(order_folder, destination)
            self.logger(f"Processed order moved back: {os.path.basename(order_folder)}")
    
    def process_pcr_folder(self, folder):
        """Process a PCR folder"""
        self.logger(f"Processing PCR folder: {os.path.basename(folder)}")
        
        # Check if already processed
        was_mseqed, has_braces, has_ab1_files = self.check_order_status(folder)
        
        if not was_mseqed and not has_braces and has_ab1_files:
            self.ui_automation.process_folder(folder)
            self.logger(f"mSeq completed: {os.path.basename(folder)}")
        else:
            self.logger(f"mSeq NOT completed: {os.path.basename(folder)}")
    
    # Additional method for FolderProcessor class

    def _get_expected_file_count(self, order_number):
        """Get expected number of files for an order based on the order key"""
        # Load the order key file
        order_key = self.file_dao.load_order_key(self.config.KEY_FILE_PATH)
        if order_key is None:
            self.logger(f"Warning: Could not load order key file, unable to verify count for order {order_number}")
            return 0
        
        # Count matching entries for this order number
        count = 0
        for row in order_key:
            if str(row[0]) == str(order_number):
                count += 1
        
        return count

    def _get_order_sample_names(self, order_number):
        """Get sample names for an order"""
        order_key = self.file_dao.load_order_key(self.config.KEY_FILE_PATH)
        if order_key is None:
            return []
        
        sample_names = []
        import numpy as np
        
        # Find indices of rows matching this order number
        indices = np.where((order_key == str(order_number)))
        
        for i in range(len(indices[0])):
            # Get the sample name and adjust characters
            sample_name = self._adjust_abi_chars(order_key[indices[0][i], 3])
            sample_names.append(sample_name)
        
        return sample_names

    def _adjust_abi_chars(self, file_name):
        """Adjust characters in file name to match ABI naming conventions"""
        # Create translation table
        translation_table = str.maketrans({
            ' ': '',
            '+': '&',
            '*': '-',
            '|': '-',
            '/': '-',
            '\\': '-',
            ':': '-',
            '"': '',
            "'": '',
            '<': '-',
            '>': '-',
            '?': '',
            ',': ''
        })
        
        # Apply translation
        return file_name.translate(translation_table)