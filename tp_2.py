#tp_2.py
from helper_funcs import (
    load_data_file,
    load_working_paper,
    get_column_indexes,
    extract_tradename_uif,
    create_output_directory,
    save_working_paper,
    insert_rows,
    copy_formatting,
    unmerge_cells_in_range,
    convert_to_dataframe,
    reapply_merged_cells,
    validate_columns,
    reset_row_heights,
    get_working_paper_path_for_all_processing
)
from tp_2_1 import (
    aggregate_data_2_1,
    populate_sheet_2_1,
)
from tp_2_2 import (
    aggregate_data_2_2,
    populate_sheet_2_2,
    apply_conditional_formatting_2_2
)
from datetime import datetime


def process_files(data_file_path, working_paper_path, consultant_name, output_directory):
    """
    Main function to process the data file and update the working paper.

    This function loads the data file and the working paper template, processes the data,
    and populates the working paper with aggregated employee data. The modified 
    working paper is then saved in the specified output directory.

    Args:
        data_file_path (str): Path to the data file (Excel).
        working_paper_path (str): Path to the working paper template (Excel).
        consultant_name (str): Name of the consultant (kept for compatibility).
        output_directory (str): Directory to save the processed working paper.

    Returns:
        str: Path to the processed working paper file.

    Raises:
        FileNotFoundError: If the data file or working paper file is not found.
        KeyError: If required columns are missing in the data file.
        Exception: If any other unexpected error occurs.
    """
    try:
        # Load the data file (Excel workbook and the first sheet)
        data_wb, data_sheet = load_data_file(data_file_path)

        # Load the working paper template (no lead sheet needed)
        working_paper_wb, _ = load_working_paper(working_paper_path, sh_n=0)

        # Extract column indexes from the data sheet
        headings = get_column_indexes(data_sheet)
        
        # Convert the source data to a pandas DataFrame
        data = convert_to_dataframe(data_sheet)

        # Validate necessary columns
        required_columns_0 = ["IDNUMBER", "FIRSTNAME", "LASTNAME"]
        validate_columns(data, required_columns_0)

        # Populate the employee Sheet 1 with aggregated employee data (TP2.1)
        employee_sheet_1 = working_paper_wb.worksheets[0]  # First sheet is now TP2.1
        populate_employee_sheet_1(employee_sheet_1, data_sheet)
        
        # Populate the employee Sheet 2 with aggregated employee data (TP2.2)
        employee_sheet_2 = working_paper_wb.worksheets[1]  # Second sheet is now TP2.2
        populate_employee_sheet_2(employee_sheet_2, data_sheet)

        # Create an output directory and get the processed file path
        tradename, uif_reference = extract_tradename_uif(data_sheet, headings)
        processed_file_path = create_output_directory(output_directory, tradename, wp_n=2, uif_reference=uif_reference, data_file_path=data_file_path, template_paths=[working_paper_path])

        # Save the modified working paper
        save_working_paper(working_paper_wb, processed_file_path)

        return processed_file_path

    except FileNotFoundError as e:
        print(f"Error: File not found - {e}")
    except KeyError as e:
        print(f"Error: Missing column in data file - {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


def process_files_for_all_processing(data_file_path, working_paper_path, consultant_name, audit_working_papers_folder):
    """
    Main function to process the data file and update the working paper.
    This version is used when processing all working papers together.

    This function loads the data file and the working paper template, processes the data,
    and populates the working paper with aggregated employee data. The modified 
    working paper is then saved in the pre-created AUDIT WORKING PAPERS folder.

    Args:
        data_file_path (str): Path to the data file (Excel).
        working_paper_path (str): Path to the working paper template (Excel).
        consultant_name (str): Name of the consultant (kept for compatibility).
        audit_working_papers_folder (str): Path to the AUDIT WORKING PAPERS subfolder.

    Returns:
        str: Path to the processed working paper file.

    Raises:
        FileNotFoundError: If the data file or working paper file is not found.
        KeyError: If required columns are missing in the data file.
        Exception: If any other unexpected error occurs.
    """
    try:
        # Load the data file (Excel workbook and the first sheet)
        data_wb, data_sheet = load_data_file(data_file_path)

        # Load the working paper template (no lead sheet needed)
        working_paper_wb, _ = load_working_paper(working_paper_path, sh_n=0)

        # Extract column indexes from the data sheet
        headings = get_column_indexes(data_sheet)
        
        # Convert the source data to a pandas DataFrame
        data = convert_to_dataframe(data_sheet)

        # Validate necessary columns
        required_columns_0 = ["IDNUMBER", "FIRSTNAME", "LASTNAME"]
        validate_columns(data, required_columns_0)

        # Populate the employee Sheet 1 with aggregated employee data (TP2.1)
        employee_sheet_1 = working_paper_wb.worksheets[0]  # First sheet is now TP2.1
        populate_employee_sheet_1(employee_sheet_1, data_sheet)
        
        # Populate the employee Sheet 2 with aggregated employee data (TP2.2)
        employee_sheet_2 = working_paper_wb.worksheets[1]  # Second sheet is now TP2.2
        populate_employee_sheet_2(employee_sheet_2, data_sheet)

        # Get the processed file path in the pre-created folder structure
        tradename, uif_reference = extract_tradename_uif(data_sheet, headings)
        processed_file_path = get_working_paper_path_for_all_processing(audit_working_papers_folder, tradename, wp_n=2, uif_reference=uif_reference)

        # Save the modified working paper
        save_working_paper(working_paper_wb, processed_file_path)

        return processed_file_path

    except FileNotFoundError as e:
        print(f"Error: File not found - {e}")
    except KeyError as e:
        print(f"Error: Missing column in data file - {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


def populate_employee_sheet_1(employee_sheet_1, data_sheet):
    """
    Populate the first sheet (TP2.1) with aggregated employee data.

    This function processes the employee data from the source sheet, aggregates the data,
    and populates it into the employee sheet. It handles row insertion and formatting.

    Args:
        employee_sheet_1 (openpyxl.worksheet.worksheet.Worksheet): The sheet object to populate.
        data_sheet (openpyxl.worksheet.worksheet.Worksheet): The sheet object containing the source data.

    Raises:
        KeyError: If a required column is missing in the data.
        Exception: If an unexpected error occurs during processing.
    """
    try:
        # 1. Convert the source data to a pandas DataFrame
        data = convert_to_dataframe(data_sheet)

        # 2. Validate the presence of required columns in the DataFrame
        required_columns_1 = [
            "IDNUMBER", "FIRSTNAME", "LASTNAME", "EMPLOYMENTSTARTDATE", "TERMINATIONDATE", 
            "BANK_PAY_AMOUNT", "LEAVE_INCOME", "MONTHLY_SALARY"
        ]
        validate_columns(data, required_columns_1)

        # 3. Aggregate the data based on specific criteria 
        aggregated = aggregate_data_2_1(data)

        # 4. Prepare for row insertion
        num_rows_to_add = len(aggregated) 
        merged_cells_to_restore = unmerge_cells_in_range(employee_sheet_1, start_row=15, end_row=26)

        # 5. Insert new rows into the target sheet
        insert_rows(employee_sheet_1, num_rows_to_add, insert_start_row=13) 
        start_row_1 = 13

        # 6. Copy formatting from existing rows to the newly inserted rows
        copy_formatting(employee_sheet_1, start_row_1, num_rows_to_add, source_cell_n=12) 

        # 7. Populate the sheet with the aggregated data
        populate_sheet_2_1(employee_sheet_1, aggregated)

        # 8. Restore any merged cells that were temporarily unmerged
        reapply_merged_cells(employee_sheet_1, merged_cells_to_restore, num_rows_to_add)

        # 9. Hide the reference row used for copying formatting
        reset_row_heights(employee_sheet_1, reference_row=12, target_rows=range(17, 19), hide_reference_row=True)

    except KeyError as e:
        print(f"Error: Missing column during employee sheet population - {e}")
    except Exception as e:
        print(f"An unexpected error occurred while populating the employee sheet: {e}")

def populate_employee_sheet_2(employee_sheet_2, data_sheet):
    """
    Populate Employee Sheet 2 (TP2.2) with aggregated employee data.

    This function takes data from the provided source sheet, aggregates it, and populates 
    it into Employee Sheet 2. It handles row insertion and formatting.

    Args:
        employee_sheet_2 (openpyxl.worksheet.worksheet.Worksheet): The sheet to populate.
        data_sheet (openpyxl.worksheet.worksheet.Worksheet): The sheet containing source employee data.

    Raises:
        KeyError: If a required column is missing in the source data.
        Exception: If an unexpected error occurs during the data processing.
    """
    try:
        # 1. Convert the source data to a pandas DataFrame
        data = convert_to_dataframe(data_sheet)

        # 2. Validate the presence of required columns in the data
        required_columns_2 = ["IDNUMBER", "LASTNAME", "FIRSTNAME", "EMPLOYMENTSTARTDATE"]
        validate_columns(data, required_columns_2)

        # 3. Aggregate the data based on specific criteria 
        aggregated = aggregate_data_2_2(data)

        # 4. Calculate the number of rows to add to the sheet
        num_rows_to_add = len(aggregated)
        merged_cells_to_restore = unmerge_cells_in_range(employee_sheet_2, start_row=16, end_row=27)

        # 5. Insert new rows into the target sheet
        start_row_2 = 14
        insert_rows(employee_sheet_2, num_rows_to_add, start_row_2)

        # 6. Copy formatting from existing rows to newly inserted rows
        copy_formatting(employee_sheet_2, start_row_2, num_rows_to_add, source_cell_n=13)

        # 7. Populate the sheet with aggregated data using specific column mappings
        populate_sheet_2_2(
            employee_sheet_2,
            aggregated,
            mappings={
                "A": lambda i, row: i,  # Row numbering
                "B": lambda i, row: "",  # Blank column
                "C": lambda i, row: row["IDNUMBER"],
                "D": lambda i, row: row["LASTNAME"],
                "E": lambda i, row: row["FIRSTNAME"],
                "F": lambda i, row: f"{row['FIRSTNAME'][0]}{row['LASTNAME'][0]}",  # Initials
                "G": lambda i, row: row["EMPLOYMENTSTARTDATE"]
            }
        )

        # 8. Restore any merged cells that were temporarily unmerged
        reapply_merged_cells(employee_sheet_2, merged_cells_to_restore, num_rows_to_add)

        # 9. Hide the reference row used for copying formatting
        reset_row_heights(employee_sheet_2, reference_row=13, target_rows=range(18, 20), hide_reference_row=True)

    except KeyError as e:
        print(f"Error: Missing required column in data file - {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
