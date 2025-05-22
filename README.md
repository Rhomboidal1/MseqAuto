```markdown
# MseqAuto

Automation tools for DNA sequencing workflow management

## Features

- Automated file sorting for Individual and Plate sequencing samples
- mSeq processing automation with quality control
- File compression and validation for data delivery
- Support for PCR project handling

## Installation

```bash
# Install in development mode
pip install -e .

# For regular installation
pip install .
```

## Usage

### Individual Sequencing Workflow

1. Sort files:
```bash
ind-sort
```

2. Run mSeq:
```bash
ind-mseq
```

3. Zip files:
```bash
ind-zip
```

### Plate Sequencing Workflow

1. Sort files:
```bash
plate-sort
```

2. Run mSeq:
```bash
plate-mseq
```

3. Zip files:
```bash
plate-zip
```

## Project Structure

```
mseqauto/
├── __init__.py         # Package initialization
├── config.py           # Configuration settings
├── core/               # Core functionality
│   ├── __init__.py
│   ├── file_system_dao.py
│   ├── folder_processor.py
│   ├── os_compatibility.py
│   └── ui_automation.py
├── scripts/            # Executable scripts
│   ├── __init__.py
│   ├── ind_sort_files.py
│   ├── ind_auto_mseq.py
│   ├── ind_zip_files.py
│   ├── plate_sort_files.py
│   ├── plate_sort_complete.py
│   ├── plate_auto_mseq.py
│   ├── plate_zip_files.py
│   ├── full_plasmid_zip_files.py
│   └── ...
├── utils/              # Utility functions
│   ├── __init__.py
│   ├── logger.py
│   └── excel_dao.py
└── legacy/             # Legacy code for reference
    └── scripts/
```

## Requirements

- Python 3.6+
- Windows OS (for UI automation)
- mSeq software installed

## Development

```bash
# Clone the repository
git clone https://github.com/Rhomboidal1/MseqAuto.git
cd MseqAuto

# Install development dependencies
pip install -e ".[dev]"

# Run tests
pytest
```

## License

[MIT License](LICENSE)
```

This format should display correctly in your repository with proper Markdown formatting and code blocks. It provides a clean overview of your project, its structure, and how to use it.