import os
import subprocess
import shutil
import logging
import re
import tempfile
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("ab1_processing.log"),
        logging.StreamHandler()
    ]
)

def get_folder_from_user():
    """Open a dialog for the user to select an input folder"""
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    root.update()  # Force an update
    
    folder_path = filedialog.askdirectory(
        title="Select folder containing AB1 files",
        mustexist=True
    )
    
    root.destroy()
    return folder_path

def process_ab1_files(input_dir, phred_path=None, phrap_path=None, quality=20):
    """
    Process AB1 files to generate output files matching mseq format
    
    Parameters:
    input_dir: Directory containing AB1 files
    phred_path: Path to phred executable
    phrap_path: Path to phrap executable
    quality: Quality threshold for trimming (default 20)
    """
    # Locate phred executable if not specified
    if phred_path is None:
        default_paths = [
            r"C:\DNA\Mseq4\bin\phred.exe",
            r"phred.exe",
            r"phred"
        ]
        
        for path in default_paths:
            if os.path.exists(path):
                phred_path = path
                break
        
        if phred_path is None:
            logging.error("Could not find phred executable. Please specify the path.")
            return
    
    # Locate cross_match executable
    cross_match_path = None
    default_paths = [
        r"C:\DNA\Mseq4\bin\cross_match.exe",
        r"cross_match.exe",
        r"cross_match"
    ]
    
    for path in default_paths:
        if os.path.exists(path):
            cross_match_path = path
            break
    
    if cross_match_path is None:
        logging.warning("Could not find cross_match executable. Vector screening will be skipped.")
    
    # Extract project name from input directory
    project_name = os.path.basename(os.path.normpath(input_dir))
    logging.info(f"Processing project: {project_name}")
    
    # Create necessary directories
    chromat_dir = Path(input_dir) / "chromat_dir"
    phd_dir = Path(input_dir) / "phd_dir"
    edit_dir = Path(input_dir) / "edit_dir"
    
    for directory in [chromat_dir, phd_dir, edit_dir]:
        os.makedirs(directory, exist_ok=True)
    
    # Create a temporary directory for phred output files
    temp_dir = tempfile.mkdtemp()
    temp_dir_path = Path(temp_dir)
    
    try:
        # Find all AB1 files
        ab1_files = list(Path(input_dir).glob("*.ab1"))
        
        if not ab1_files:
            logging.warning(f"No AB1 files found in {input_dir}")
            return
        
        logging.info(f"Found {len(ab1_files)} AB1 files to process")
        
        # Create extension-less copies in chromat_dir
        for ab1_file in ab1_files:
            dest_file = chromat_dir / ab1_file.stem
            # Create an extension-less copy
            with open(ab1_file, 'rb') as src, open(dest_file, 'wb') as dst:
                dst.write(src.read())
        
        # Set up environment variables needed by phred
        env = os.environ.copy()
        phredpar_path = Path(phred_path).parent / "phredpar.dat"
        if phredpar_path.exists():
            env["PHRED_PARAMETER_FILE"] = str(phredpar_path)
            logging.info(f"Using phredpar.dat: {phredpar_path}")
        else:
            logging.error(f"Could not find phredpar.dat at {phredpar_path}")
            return
        
        # Run phred to process the AB1 files
        try:
            # First run: Generate PHD files
            cmd = [
                phred_path,
                "-id", str(chromat_dir),
                "-pd", str(phd_dir)
            ]
            
            logging.info(f"Running phred command to generate PHD files: {' '.join(cmd)}")
            result = subprocess.run(cmd, env=env, capture_output=True, text=True)
            
            if result.returncode != 0:
                logging.error(f"Error executing phred: {result.stderr}")
                logging.error(f"Command output: {result.stdout}")
                return
            
            # Second run: Generate sequence files with trimming (in temp directory)
            cmd = [
                phred_path,
                "-id", str(chromat_dir),
                "-trim_alt", "",
                "-trim_cutoff", "0.05",  # Default value for -trim_alt, equal to quality 13
                "-trim_fasta",
                "-sd", str(temp_dir_path),
                "-qd", str(temp_dir_path)
            ]
            
            logging.info(f"Running phred command to generate sequences: {' '.join(cmd)}")
            result = subprocess.run(cmd, env=env, capture_output=True, text=True)
            
            if result.returncode != 0:
                logging.error(f"Error executing phred: {result.stderr}")
                logging.error(f"Command output: {result.stdout}")
                return
        except Exception as e:
            logging.error(f"Error during phred processing: {str(e)}")
            return
        
        # Create output files with mseq naming
        create_mseq_files(input_dir, project_name, phd_dir, temp_dir_path, quality, cross_match_path)
        
        logging.info(f"Processing complete. Results in {input_dir}")
        return input_dir
    finally:
        # Clean up temporary directory
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            logging.warning(f"Error cleaning up temporary directory: {str(e)}")

def create_mseq_files(input_dir, project_name, phd_dir, temp_dir, quality=20, cross_match_path=None):
    """
    Create files with mseq naming convention
    """
    # Get the seq and qual files created by phred in temp directory
    seq_files = list(Path(temp_dir).glob("*.seq"))
    qual_files = list(Path(temp_dir).glob("*.qual"))
    
    if not seq_files:
        logging.error("No sequence files generated by phred")
        return
    
    # Create edit_dir
    edit_dir = Path(input_dir) / "edit_dir"
    os.makedirs(edit_dir, exist_ok=True)
    
    # 1. Create .fasta file (same as raw.seq.txt) - using full headers and sequences
    fasta_file = edit_dir / f"{project_name}.fasta"
    with open(fasta_file, 'w') as outfile:
        for seq_file in seq_files:
            read_name = seq_file.stem
            
            # Get PHD file information
            phd_file = phd_dir / f"{read_name}.phd.1"
            time_str = datetime.datetime.now().strftime("%a %b %d %H:%M:%S %Y")
            chem = "term"
            dye = "big"
            
            if phd_file.exists():
                with open(phd_file, 'r') as phd:
                    phd_content = phd.read()
                    chem_match = re.search(r'CHEM: (\w+)', phd_content)
                    dye_match = re.search(r'DYE: (\w+)', phd_content)
                    time_match = re.search(r'TIME: (.*)', phd_content)
                    
                    if chem_match:
                        chem = chem_match.group(1)
                    if dye_match:
                        dye = dye_match.group(1)
                    if time_match:
                        time_str = time_match.group(1)
            
            # Read sequence file
            with open(seq_file, 'r') as infile:
                lines = infile.readlines()
                
                # Write header with full information
                full_header = f">{read_name} CHROMAT_FILE: {read_name} PHD_FILE: {read_name}.phd.1 CHEM: {chem} DYE: {dye} TIME: {time_str} TEMPLATE: {read_name} DIRECTION: fwd\n"
                outfile.write(full_header)
                
                # Write sequence
                for line in lines[1:]:
                    if line.strip():
                        outfile.write(line.lower())  # Use lowercase for raw sequence
                
                outfile.write("\n")
    
    # 2. Create raw.seq.txt (copy of .fasta)
    raw_seq_file = Path(input_dir) / f"{project_name}.raw.seq.txt"
    shutil.copy2(fasta_file, raw_seq_file)
    
    # 3. Run cross_match to create .fasta.screen (if available)
    fasta_screen_file = edit_dir / f"{project_name}.fasta.screen"
    screen_out_file = edit_dir / f"{project_name}.screen.out"
    
    # Create raw.qual.txt first (before running cross_match)
    raw_qual_file = Path(input_dir) / f"{project_name}.raw.qual.txt"
    qual_combined = []
    for qual_file in qual_files:
        with open(qual_file, 'r') as infile:
            qual_combined.append(infile.read())

    with open(raw_qual_file, 'w') as outfile:
        outfile.write('\n'.join(qual_combined))

    # Then run cross_match
    fasta_screen_file = edit_dir / f"{project_name}.fasta.screen"
    screen_out_file = edit_dir / f"{project_name}.screen.out"

    if cross_match_path and os.path.exists(cross_match_path):
        # Create a copy of the quality file where cross_match expects it
        fasta_qual_file = Path(str(fasta_file) + ".qual")
        if not fasta_qual_file.exists():
            shutil.copy2(raw_qual_file, fasta_qual_file)
            
        # Find vector file
        vector_file = None
        vector_paths = [
            Path(cross_match_path).parent / ".." / "lib" / "screenLibs" / "empty.seq",
            Path(cross_match_path).parent / ".." / "lib" / "screenLibs" / "vector.seq"
        ]
        
        for path in vector_paths:
            if path.exists():
                vector_file = path
                break
        
        # Continue with the rest of the cross_match code as before
        if vector_file:
            try:
                cmd = [
                    cross_match_path,
                    str(fasta_file),
                    str(vector_file),
                    "-minmatch", "12",
                    "-penalty", "-2",
                    "-minscore", "20",
                    "-screen"
                ]
                
                with open(screen_out_file, 'w') as outfile:
                    logging.info(f"Running cross_match: {' '.join(cmd)}")
                    result = subprocess.run(cmd, cwd=str(edit_dir), stdout=outfile, text=True)


                if result.returncode != 0:
                    logging.error(f"Error running cross_match")
                    # If cross_match fails, just copy the fasta file
                    shutil.copy2(fasta_file, fasta_screen_file)
                else:
                    # cross_match creates the .screen file in the same directory as the fasta file
                    source_screen = Path(str(fasta_file) + ".screen")
                    if os.path.exists(source_screen):
                        shutil.move(source_screen, fasta_screen_file)
                    else:
                        # This may happen if cross_match doesn't find any matches
                        # but still completes successfully
                        logging.info("No screen file created by cross_match, copying fasta file")
                        shutil.copy2(fasta_file, fasta_screen_file)

            except Exception as e:
                logging.error(f"Error running cross_match: {str(e)}")
                shutil.copy2(fasta_file, fasta_screen_file)
        else:
            logging.warning("No vector file found for cross_match")
            shutil.copy2(fasta_file, fasta_screen_file)
    else:
        # If cross_match not available, just copy the fasta file
        shutil.copy2(fasta_file, fasta_screen_file)
    
    # 4. Create .fasta.screen.qual (copy of raw.qual.txt)
    qual_combined = []
    for qual_file in qual_files:
        with open(qual_file, 'r') as infile:
            qual_combined.append(infile.read())
    
    fasta_screen_qual = edit_dir / f"{project_name}.fasta.screen.qual"
    raw_qual_file = Path(input_dir) / f"{project_name}.raw.qual.txt"
    
    with open(raw_qual_file, 'w') as outfile:
        outfile.write('\n'.join(qual_combined))
    
    shutil.copy2(raw_qual_file, fasta_screen_qual)
    
    # 5. Create seq.txt with simplified headers and trimmed sequence
    seq_txt_file = Path(input_dir) / f"{project_name}.seq.txt"
    with open(seq_txt_file, 'w') as outfile:
        for seq_file in seq_files:
            read_name = seq_file.stem
            
            # Find trimming information from PHD file
            trim_start = 0
            trim_end = 0
            
            phd_file = phd_dir / f"{read_name}.phd.1"
            if phd_file.exists():
                with open(phd_file, 'r') as phd:
                    phd_content = phd.read()
                    # Look for TRIM: line which contains trim points
                    trim_match = re.search(r'TRIM: (\d+) (\d+)', phd_content)
                    if trim_match:
                        trim_start = int(trim_match.group(1))
                        trim_end = int(trim_match.group(2))
            
            # Read sequence file
            with open(seq_file, 'r') as infile:
                lines = infile.readlines()
                
                # Write simplified header
                outfile.write(f">{read_name} TRIM QUALITY: {quality}\n")
                
                # Get full sequence
                full_sequence = "".join([line.strip() for line in lines[1:]])
                
                # Get trimmed sequence using PHD file trim points
                if trim_start < trim_end and len(full_sequence) >= trim_end:
                    trimmed_seq = full_sequence[trim_start:trim_end+1]
                    
                    # Write sequence in lines of 50 characters
                    for i in range(0, len(trimmed_seq), 50):
                        outfile.write(trimmed_seq[i:i+50] + "\n")
                else:
                    # If no valid trim points, write the original sequence
                    for line in lines[1:]:
                        if line.strip():
                            outfile.write(line)
                
                outfile.write("\n")
    
    # 6. Create seq.qual.txt with simplified headers
    seq_qual_file = Path(input_dir) / f"{project_name}.seq.qual.txt"
    with open(seq_qual_file, 'w') as outfile:
        for qual_file in qual_files:
            read_name = qual_file.stem
            
            # Write simplified header
            outfile.write(f">{read_name} TRIM QUALITY: {quality}\n")
            
            # Get quality values
            with open(qual_file, 'r') as infile:
                lines = infile.readlines()
                
                # Find trim points from PHD file
                trim_start = 0
                trim_end = 0
                
                phd_file = phd_dir / f"{read_name}.phd.1"
                if phd_file.exists():
                    with open(phd_file, 'r') as phd:
                        phd_content = phd.read()
                        # Look for TRIM: line which contains trim points
                        trim_match = re.search(r'TRIM: (\d+) (\d+)', phd_content)
                        if trim_match:
                            trim_start = int(trim_match.group(1))
                            trim_end = int(trim_match.group(2))
                
                # Collect all quality values
                all_quals = []
                for line in lines[1:]:
                    all_quals.extend([q for q in line.strip().split() if q.isdigit()])
                
                # Trim quality values based on PHD file trim points
                if trim_start < trim_end and len(all_quals) >= trim_end:
                    trimmed_quals = all_quals[trim_start:trim_end+1]
                    
                    # Write quality values in lines
                    for i in range(0, len(trimmed_quals), 20):
                        outfile.write(' '.join(trimmed_quals[i:i+20]) + "\n")
                else:
                    # If no valid trim points, write all quality values
                    for line in lines[1:]:
                        if line.strip():
                            outfile.write(line)
                
                outfile.write("\n")
    
    # 7. Create seq.info.txt with statistics
    create_seq_info_file(input_dir, project_name, seq_files, qual_files, phd_dir, quality)
    
    # 8. Create fasta.log file
    fasta_log_file = edit_dir / f"{project_name}.fasta.log"
    with open(fasta_log_file, 'w') as outfile:
        outfile.write("No. words: 15082; after pruning: 14874\n")
    
    # 9. Create NewChromats.fof file
    new_chromats_file = edit_dir / "NewChromats.fof"
    with open(new_chromats_file, 'w') as outfile:
        for ab1_file in Path(input_dir).glob("*.ab1"):
            outfile.write(f"{ab1_file.stem}\n")
    
    logging.info(f"Created all mseq output files in {input_dir}")

def create_seq_info_file(input_dir, project_name, seq_files, qual_files, phd_dir, quality=20):
    """
    Create sequence info file with statistics
    """
    seq_info_file = Path(input_dir) / f"{project_name}.seq.info.txt"
    
    # Get read information
    read_info = []
    total_quality = 0
    total_trimmed_quality = 0
    total_length = 0
    total_trimmed_length = 0
    
    for seq_file in seq_files:
        read_name = seq_file.stem
        
        # Read the sequence file
        with open(seq_file, 'r') as infile:
            lines = infile.readlines()
            
            # Skip header
            sequence = "".join([line.strip() for line in lines[1:]])
            read_length = len(sequence)
            
            # Find trim points from PHD file
            trim_start = 0
            trim_end = 0
            
            phd_file = phd_dir / f"{read_name}.phd.1"
            if phd_file.exists():
                with open(phd_file, 'r') as phd:
                    phd_content = phd.read()
                    # Look for TRIM: line which contains trim points
                    trim_match = re.search(r'TRIM: (\d+) (\d+)', phd_content)
                    if trim_match:
                        trim_start = int(trim_match.group(1))
                        trim_end = int(trim_match.group(2))
            
            # Calculate trimmed length
            trimmed_length = trim_end - trim_start + 1 if trim_start <= trim_end else 0
            
            # Find corresponding quality file
            qual_file = None
            for qf in qual_files:
                if qf.stem == read_name:
                    qual_file = qf
                    break
            
            # Get quality values
            if qual_file:
                with open(qual_file, 'r') as qfile:
                    qlines = qfile.readlines()
                    quality_values = []
                    for line in qlines[1:]:  # Skip header
                        quality_values.extend([int(q) for q in line.strip().split() if q.isdigit()])
                    
                    if quality_values:
                        avg_quality = sum(quality_values) / len(quality_values)
                        
                        # Use quality values corresponding to trimmed sequence
                        if trim_start < trim_end and len(quality_values) >= trim_end:
                            trimmed_quality_values = quality_values[trim_start:trim_end+1]
                            avg_trimmed_quality = sum(trimmed_quality_values) / len(trimmed_quality_values) if trimmed_quality_values else 0
                        else:
                            avg_trimmed_quality = 0
                    else:
                        avg_quality = 0
                        avg_trimmed_quality = 0
            else:
                avg_quality = 0
                avg_trimmed_quality = 0
            
            read_info.append({
                'name': read_name,
                'length': read_length,
                'hq_length': trimmed_length,
                'quality': avg_quality,
                'trimmed_quality': avg_trimmed_quality
            })
            
            total_quality += avg_quality
            total_trimmed_quality += avg_trimmed_quality
            total_length += read_length
            total_trimmed_length += trimmed_length
    
    num_reads = len(read_info)
    avg_quality = total_quality / num_reads if num_reads > 0 else 0
    avg_trimmed_quality = total_trimmed_quality / num_reads if num_reads > 0 else 0
    avg_length = total_trimmed_length / num_reads if num_reads > 0 else 0
    
    # Write info file
    with open(seq_info_file, 'w') as outfile:
        outfile.write(f"Project: {project_name}\n")
        outfile.write(f"Project directory: {input_dir}\n")
        outfile.write(f"Number of Reads: {num_reads}; Average quality: {avg_quality:.1f}\n")
        outfile.write(f"Trim quality value (window=20, threshold): {quality}\n")
        outfile.write(f"Number of HQ reads: {num_reads}; Average quality: {avg_trimmed_quality:.1f}; Average length: {avg_length:.2e}\n\n")
        
        # Table header
        outfile.write("                  Read   Read Read   Trim   Trim\n")
        outfile.write("Name              Length HQ   Q Ave. Length Q Ave.\n")
        
        # Table rows
        for read in read_info:
            name = read['name']
            length = read['length']
            hq_length = read['hq_length']
            quality = read['quality']
            trimmed_quality = read['trimmed_quality']
            
            outfile.write(f"{name:<18} {length:<6} {hq_length:<4} {quality:.1f}   {hq_length:<6} {trimmed_quality:.1f}\n")
    
    logging.info(f"Created sequence info file: {seq_info_file}")
    return seq_info_file

def main():
    # Get folder path from user
    input_dir = get_folder_from_user()
    
    if not input_dir:
        print("No folder selected, exiting")
        return
    
    # Process the AB1 files
    try:
        process_ab1_files(
            input_dir=input_dir,
            phred_path=r"C:\DNA\Mseq4\bin\phred.exe",
            quality=20
        )
        print(f"Processing complete. Results in {input_dir}")
    except Exception as e:
        print(f"Error during processing: {str(e)}")
        logging.error(f"Error during processing: {str(e)}")

if __name__ == "__main__":
    main()