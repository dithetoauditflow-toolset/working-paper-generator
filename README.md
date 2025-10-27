# Working Paper Generator

## Overview

The **Working Paper Generator** is a desktop application built with Tkinter that automates the creation of UIF TERS audit working papers (TP.1, TP.2, TP.3, and TP.4). It processes Excel data files using macro-enabled templates to generate standardized, populated working papers with structured output folders and comprehensive error handling.

## Purpose

Automate the generation of audit working papers by processing UIF data files through standardized templates, eliminating manual data entry and ensuring consistency across all audit documentation.

## Key Features

- **Graphical User Interface**: Tkinter-based GUI with centered layout, custom branding, and intuitive navigation
- **Single & Batch Processing**: Select single or multiple Excel data files for processing
- **Consultant Information Capture**: Prompts for consultant name to personalize working papers
- **Template Management**: Automatically retrieves and validates `.xlsm` template files from `TEMPLATES` directory
- **Working Paper Generation**: Creates TP.1 (Company Info), TP.2 (Claims Analysis), TP.3 (Payments Analysis), and TP.4 (Reconciliation)
- **Individual or Batch Generation**: Generate specific working papers or all four at once
- **Progress Tracking**: Real-time progress bar during generation process
- **Results Display**: Table showing processing status and time taken for each file
- **Output Management**: User-selectable output directory with organized folder structure
- **Reset Functionality**: Clear selections and start new processing session
- **Comprehensive Error Handling**: Informative error messages for missing templates, data format issues, and processing errors

## Business Benefits

- **Time Savings**: Reduce working paper generation time from hours to minutes
- **Consistency**: Ensure all working papers follow standardized templates and formatting
- **Error Reduction**: Eliminate manual data entry errors and formula mistakes
- **Audit Quality**: Maintain high-quality, professional working papers for all engagements
- **Scalability**: Process multiple companies quickly during peak audit periods
- **Compliance**: Ensure all working papers meet UIF audit standards

## Tech Stack

- **Python**: 3.11+
- **GUI Framework**: Tkinter (built-in), ttk (themed widgets)
- **Data Processing**: Pandas 1.5.0+, OpenPyXL 3.1.0+
- **Template Engine**: Excel macro-enabled workbooks (.xlsm)

## Inputs

- **Data Files**: Excel files (`.xlsx`) containing UIF TERS employer data
  - Required columns: UIF reference, company name, employee details, claims data, payments data
- **Templates**: Four macro-enabled Excel templates (`.xlsm`) in `TEMPLATES` directory
  - TP.1 template: Company information and background
  - TP.2 template: Claims analysis and testing
  - TP.3 template: Payments analysis and verification
  - TP.4 template: Reconciliation and conclusions
- **Consultant Name**: User-provided name for working paper attribution

## Outputs

- **Generated Working Papers**: Populated Excel files (TP.1, TP.2, TP.3, TP.4) in user-selected output directory
- **Organized Folders**: Structured output with company-specific subfolders
- **Processing Results**: Summary table showing success/failure status and processing time

## How It Works

1. **Application Launch**: User runs `main.py` or `ui.py` to launch Tkinter GUI
2. **File Selection**: User selects single or multiple Excel data files via file dialog
3. **Template Validation**: Application verifies four `.xlsm` templates exist in `TEMPLATES` directory
4. **Consultant Input**: Dialog prompts user to enter consultant name
5. **Output Selection**: User selects destination directory for generated working papers
6. **Working Paper Choice**: User selects individual working papers (TP.1-4) or "Generate All"
7. **Data Extraction**: Application reads data files and extracts company information using `helper_funcs.py`
8. **Template Population**: Each TP script (`tp_1.py`, `tp_2.py`, `tp_3.py`, `tp_4.py`) opens template, populates data, and saves output
9. **Progress Display**: Progress bar updates as each working paper is generated
10. **Results Summary**: Table displays processing results with status and time taken
11. **Folder Organization**: Output files organized by company name and working paper type

## Installation & Setup

### Prerequisites
- Python 3.11 or higher
- Excel with macro support (for template editing)
- UIF data files in Excel format

### Installation Steps

1. **Navigate to Tool Directory**:
   ```bash
   cd working_paper_generator
   ```

2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Prepare Templates**:
   - Ensure `TEMPLATES` directory exists
   - Place four `.xlsm` template files (TP.1, TP.2, TP.3, TP.4) in `TEMPLATES` directory
   - Verify templates have correct cell references and formulas

4. **Prepare UI Assets** (Optional):
   - Place `icon.ico` in `img/` directory for window icon
   - Place `logo.png` in `img/` directory for application branding

5. **Run Application**:
   ```bash
   python main.py
   ```
   or
   ```bash
   python ui.py
   ```

## Configuration

### Template Structure
- **Location**: `TEMPLATES/` directory
- **Format**: Excel macro-enabled workbooks (`.xlsm`)
- **Required Templates**:
  - Template 1: TP.1 (Company Information)
  - Template 2: TP.2 (Claims Analysis)
  - Template 3: TP.3 (Payments Analysis)
  - Template 4: TP.4 (Reconciliation)
- **Cell References**: Templates must have predefined cell references for data population

### Data File Requirements
- **Format**: Excel workbook (`.xlsx`)
- **Required Columns**:
  - UIF reference number
  - Company/trade name
  - Employee details (names, ID numbers, positions)
  - Claims data (periods, amounts, categories)
  - Payments data (dates, amounts, bank details)

### Helper Functions (`helper_funcs.py`)
- **get_company_info()**: Extracts company information from data files
- Customize extraction logic to match your data file structure

## Directory Structure

```
working_paper_generator/
├── main.py                              # Main entry point
├── ui.py                                # Tkinter GUI interface
├── cli_ui.py                            # Command-line interface (alternative)
├── requirements.txt                      # Python dependencies
├── README.md                            # This documentation
├── TEMPLATES/                           # Template files (gitignored)
│   ├── TP1_Template.xlsm
│   ├── TP2_Template.xlsm
│   ├── TP3_Template.xlsm
│   └── TP4_Template.xlsm
├── img/                                 # UI assets
│   ├── icon.ico
│   └── logo.png
├── tp_1.py                              # TP.1 generation script
├── tp_2.py                              # TP.2 generation script
├── tp_2_1.py                            # TP.2 sub-schedule 1
├── tp_2_2.py                            # TP.2 sub-schedule 2
├── tp_3.py                              # TP.3 generation script
├── tp_3_1.py                            # TP.3 sub-schedule 1
├── tp_3_2.py                            # TP.3 sub-schedule 2
├── tp_3_3.py                            # TP.3 sub-schedule 3
├── tp_4.py                              # TP.4 generation script
└── helper_funcs.py                      # Utility functions
```

## Usage Workflow

1. **Launch Application**: Run `python main.py`
2. **Select Data Files**: Click "Select File(s)" and choose Excel data files
3. **Review Selection**: Verify selected files displayed in application
4. **Enter Consultant Name**: Provide name when prompted
5. **Select Output Directory**: Choose destination folder for generated working papers
6. **Choose Working Papers**: Select individual TPs or "Generate All"
7. **Monitor Progress**: Watch progress bar as working papers are generated
8. **Review Results**: Check results table for success/failure status
9. **Access Output**: Navigate to output directory to view generated working papers
10. **Reset (Optional)**: Click "Reset" to clear selections and process new files

## Troubleshooting

### Template Issues
- **Error**: "Templates not found"
  - **Solution**: Verify four `.xlsm` files exist in `TEMPLATES` directory
- **Error**: "Template validation failed"
  - **Solution**: Ensure at least four template files are present
- **Error**: "Template corrupted"
  - **Solution**: Re-download or restore template files from backup

### Data File Issues
- **Error**: "Failed to read data file"
  - **Solution**: Ensure file is valid Excel format (`.xlsx`) and not corrupted
- **Error**: "Missing required columns"
  - **Solution**: Verify data file has all required columns (UIF ref, company name, etc.)
- **Error**: "Data format incorrect"
  - **Solution**: Check data types match expected formats (numbers, dates, text)

### Processing Issues
- **Error**: "Working paper generation failed"
  - **Solution**: Check template cell references match data extraction logic
- **Error**: "Output file already exists"
  - **Solution**: Delete existing output files or choose different output directory
- **Error**: "Permission denied"
  - **Solution**: Ensure write permissions in output directory and close any open Excel files

### UI Issues
- **Error**: "Icon/logo not found"
  - **Solution**: Place `icon.ico` and `logo.png` in `img/` directory (non-critical)
- **Error**: "Window not displaying correctly"
  - **Solution**: Check screen resolution and Tkinter compatibility

## Best Practices

- **Template Maintenance**: Keep templates updated and version-controlled
- **Data Validation**: Verify data files are complete before processing
- **Batch Processing**: Process multiple companies in one session for efficiency
- **Output Organization**: Use clear, consistent naming for output directories
- **Template Backup**: Maintain backups of original templates before customization
- **Testing**: Test with sample data before processing production files

## Customization

- **Templates**: Modify `.xlsm` files in `TEMPLATES` directory to customize output format
- **Processing Logic**: Adapt `tp_1.py`, `tp_2.py`, `tp_3.py`, `tp_4.py` for specific requirements
- **UI Elements**: Modify `ui.py` to change appearance, layout, and branding
- **Company Info Extraction**: Update `get_company_info()` in `helper_funcs.py` to match data structure

## Version Information

- **Version**: WP2 1.4
- **Last Updated**: December 2024
- **Developed By**: Seward Mupereri for Resumes By Ropa Pty Ltd
- **License**: Proprietary
