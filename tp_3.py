#tp_3.py
from helper_funcs import (
    load_data_file,
    load_working_paper,
    get_column_indexes,
    extract_tradename_uif,
    create_output_directory,
    save_working_paper,
    insert_rows,
    copy_formatting,
    reset_row_heights,
    unmerge_cells_in_range,
    convert_to_dataframe,
    reapply_merged_cells,
    validate_columns,
    apply_conditional_formatting_general,
    column_index_to_letter,
    column_letter_to_index,
    update_formulas_after_row_insertion,
    adjust_formula_references,
    get_working_paper_path_for_all_processing
)
from tp_3_1 import (
    aggregate_data_3_1,
    populate_sheet_3_1
)
from tp_3_2 import (
    aggregate_data_3_2,
    populate_sheet_3_2,
    adjust_column_visibility,
    replicate_hidden_columns
)
from tp_3_3 import (
    aggregate_data_3_3,
    populate_sheet_3_3
)
from datetime import datetime

def populate_working_paper(lead_sheet, tradename, uif_reference, periods_str, current_date, consultant_name):
    """
    Populates the working paper with extracted company details.

    This function updates specific cells in the provided lead sheet of the working paper
    with the extracted data, including tradename, UIF reference, shutdown periods, current
    date, and the consultant's name. It safely handles merged cells by unmerging them
    before writing and then remerging them after.

    Args:
        lead_sheet (openpyxl.worksheet.worksheet.Worksheet): The worksheet object to populate with data.
        tradename (str): The tradename to be inserted into the working paper.
        uif_reference (str): The UIF reference to be inserted into the working paper.
        periods_str (str): The shutdown periods to be inserted into the working paper.
        current_date (str): The current date to be inserted into the working paper.
        consultant_name (str): The name of the consultant to be inserted into the working paper.

    Returns:
        None
    """
    try:
        # First, unmerge any cells in the company info area (rows 1-4, columns B and E)
        # This covers the range where company info is typically inserted
        merged_cells_to_restore = unmerge_cells_in_range(lead_sheet, start_row=1, end_row=4)
        
        # Insert the extracted data into specific cells in the working paper
        lead_sheet["B1"].value = tradename
        lead_sheet["B2"].value = uif_reference
        lead_sheet["B4"].value = periods_str
        lead_sheet["E3"].value = current_date
        lead_sheet["E1"].value = consultant_name  # Insert the consultant's name into E1
        
        # Reapply any merged cells that were temporarily unmerged
        # Note: We pass 0 as num_rows_added since we're not adding rows, just unmerging for writing
        reapply_merged_cells(lead_sheet, merged_cells_to_restore, 0)
        
        print(f"DEBUG: Successfully populated company info - Company: {tradename}, UIF: {uif_reference}")
        
    except Exception as e:
        print(f"ERROR in populate_working_paper: {e}")
        print(f"ERROR DETAILS - Function: populate_working_paper, Error: {type(e).__name__}")
        raise

def process_files(data_file_path, working_paper_path, consultant_name, output_directory):
    """
    Main function to process the data file and update the working paper with extracted information.
    
    This function:
    - Loads the data and working paper files.
    - Extracts necessary details such as tradename, UIF reference, shutdown periods, and current date.
    - Populates payment sheets with aggregated payment data for different categories (TP3.1, TP3.2, TP3.3).
    - Creates an output directory for saving the processed file.
    - Saves the updated working paper to the specified output directory.

    Args:
        data_file_path (str): Path to the source data file (e.g., Excel file).
        working_paper_path (str): Path to the working paper file to be updated.
        consultant_name (str): Name of the consultant responsible for the working paper.
        output_directory (str): Path to the directory where the processed working paper will be saved.

    Returns:
        str: Path to the saved working paper after processing and updates.
    
    Raises:
        Exception: If there are any errors loading the files or extracting data.
    """
    try:
        # Load the data file (Excel workbook and the first sheet)
        data_wb, data_sheet = load_data_file(data_file_path)

        # Load the working paper template (Excel workbook and first sheet)
        working_paper_wb, first_sheet = load_working_paper(working_paper_path, sh_n=0)

        # Extract column indexes from the data sheet
        headings = get_column_indexes(data_sheet)
        
        # Extract tradename and UIF reference from the data sheet
        tradename, uif_reference = extract_tradename_uif(data_sheet, headings)

        # Use fixed string for periods
        periods_str = "Lockdown Periods"

        # Get the current date for record-keeping
        current_date = datetime.now().strftime("%Y-%m-%d")

        # Populate the working paper with company details (into the first sheet - TP3.1)
        populate_working_paper(first_sheet, tradename, uif_reference, periods_str, current_date, consultant_name)

        # Populate the payments sheets with aggregated payments data (TP3.1, TP3.2, TP3.3)
        # TP3.1 is the first sheet (index 0) - same as first_sheet
        populate_payments_sheet_1(first_sheet, data_sheet)

        # TP3.2 is the second sheet (index 1)
        payment_sheet_2 = working_paper_wb.worksheets[1]
        num_rows_to_add = populate_payments_sheet_2(payment_sheet_2, data_sheet)

        # TP3.3 is the third sheet (index 2)
        payments_sheet_3 = working_paper_wb.worksheets[2]
        populate_payments_sheet_3(payments_sheet_3, data_sheet)

        # Create an output directory and get the processed file path
        processed_file_path = create_output_directory(output_directory, tradename, wp_n=3, uif_reference=uif_reference, data_file_path=data_file_path, template_paths=[working_paper_path])

        # Save the modified working paper
        save_working_paper(working_paper_wb, processed_file_path)

        return processed_file_path

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        raise


def process_files_for_all_processing(data_file_path, working_paper_path, consultant_name, audit_working_papers_folder):
    """
    Main function to process the data file and update the working paper with extracted information.
    This version is used when processing all working papers together.
    
    This function:
    - Loads the data and working paper files.
    - Extracts necessary details such as tradename, UIF reference, shutdown periods, and current date.
    - Populates payment sheets with aggregated payment data for different categories (TP3.1, TP3.2, TP3.3).
    - Saves the updated working paper to the pre-created AUDIT WORKING PAPERS folder.

    Args:
        data_file_path (str): Path to the source data file (e.g., Excel file).
        working_paper_path (str): Path to the working paper file to be updated.
        consultant_name (str): Name of the consultant responsible for the working paper.
        audit_working_papers_folder (str): Path to the AUDIT WORKING PAPERS subfolder.

    Returns:
        str: Path to the saved working paper after processing and updates.
    
    Raises:
        Exception: If there are any errors loading the files or extracting data.
    """
    try:
        # Load the data file (Excel workbook and the first sheet)
        data_wb, data_sheet = load_data_file(data_file_path)

        # Load the working paper template (Excel workbook and first sheet)
        working_paper_wb, first_sheet = load_working_paper(working_paper_path, sh_n=0)

        # Extract column indexes from the data sheet
        headings = get_column_indexes(data_sheet)
        
        # Extract tradename and UIF reference from the data sheet
        tradename, uif_reference = extract_tradename_uif(data_sheet, headings)

        # Use fixed string for periods
        periods_str = "Lockdown Periods"

        # Get the current date for record-keeping
        current_date = datetime.now().strftime("%Y-%m-%d")

        # Populate the working paper with company details (into the first sheet - TP3.1)
        populate_working_paper(first_sheet, tradename, uif_reference, periods_str, current_date, consultant_name)

        # Populate the payments sheets with aggregated payments data (TP3.1, TP3.2, TP3.3)
        # TP3.1 is the first sheet (index 0) - same as first_sheet
        populate_payments_sheet_1(first_sheet, data_sheet)

        # TP3.2 is the second sheet (index 1)
        payment_sheet_2 = working_paper_wb.worksheets[1]
        num_rows_to_add = populate_payments_sheet_2(payment_sheet_2, data_sheet)

        # TP3.3 is the third sheet (index 2)
        payments_sheet_3 = working_paper_wb.worksheets[2]
        populate_payments_sheet_3(payments_sheet_3, data_sheet)

        # Get the processed file path in the pre-created folder structure
        processed_file_path = get_working_paper_path_for_all_processing(audit_working_papers_folder, tradename, wp_n=3, uif_reference=uif_reference)

        # Save the modified working paper
        save_working_paper(working_paper_wb, processed_file_path)

        return processed_file_path

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        raise


def populate_payments_sheet_1(payments_sheet_1, data_sheet):
    """
    Populate the payments sheet (TP3.1) with aggregated data extracted from the data sheet.

    This function performs the following tasks:
    - Converts the source data to a pandas DataFrame.
    - Validates the presence of required columns.
    - Aggregates the payment data.
    - Inserts new rows into the payments sheet.
    - Applies formatting and formulas to the populated data.
    
    Args:
        payments_sheet_1 (openpyxl.worksheet.worksheet.Worksheet): The sheet object for the payments sheet to populate.
        data_sheet (openpyxl.worksheet.worksheet.Worksheet): The source data sheet containing the raw data.

    Raises:
        KeyError: If a required column is missing in the data.
        Exception: If an unexpected error occurs during processing.
    """
    try:
        # 1. Convert the source data to a pandas DataFrame
        data = convert_to_dataframe(data_sheet)

        # 2. Validate the presence of required columns in the DataFrame
        required_columns = ["PAYMENTDATE", "PAY_REF_ITR_1", "BANK_PAY_AMOUNT"]
        validate_columns(data, required_columns)

        # 3. Process and aggregate the data
        aggregated = aggregate_data_3_1(data)

        # 4. Prepare for row insertion
        num_rows_to_add = len(aggregated)
        merged_cells_to_restore = unmerge_cells_in_range(payments_sheet_1, start_row=22, end_row=37)

        # 5. Insert new rows into the target sheet
        insert_rows(payments_sheet_1, num_rows_to_add, insert_start_row=20)
        start_row = 20

        # 6. Update existing formulas to account for inserted rows
        # REMOVED: update_formulas_after_row_insertion(payments_sheet_1, 20, num_rows_to_add)
        # Keeping original formulas as they are correct

        # 7. Copy formatting from row 19
        copy_formatting(payments_sheet_1, start_row, num_rows_to_add, source_cell_n=19)

        # 8. Populate the sheet with the aggregated data
        populate_sheet_3_1(payments_sheet_1, aggregated)

        # 9. Add SUM formulas in columns D and H
        total_row = start_row + num_rows_to_add + 2  # 3rd row after the last inserted row
        sheet_range_d = f"D{start_row}:D{start_row + num_rows_to_add - 1}"
        sheet_range_h = f"H{start_row}:H{start_row + num_rows_to_add - 1}"
        payments_sheet_1[f"D{total_row}"] = f"=SUM({sheet_range_d})"
        payments_sheet_1[f"H{total_row}"] = f"=SUM({sheet_range_h})"

        # 10. Add the difference formula in column I
        payments_sheet_1[f"I{total_row}"] = f"=D{total_row} - H{total_row}"

        # 11. Restore any merged cells that were temporarily unmerged
        reapply_merged_cells(payments_sheet_1, merged_cells_to_restore, num_rows_to_add)

        # 12. Adjust row heights for better presentation
        reset_row_heights(
            payments_sheet_1,
            reference_row=19,
            target_rows=range(26, 28 + num_rows_to_add),
            hide_reference_row=True
        )
        
        # 13. Apply conditional formatting to new rows
        columns_to_format = ['F', 'G', 'H']
        apply_conditional_formatting_general(payments_sheet_1, start_row, num_rows_to_add, columns_to_format, legend='K')

    except KeyError as e:
        print(f"Error: Missing column during payments sheet population - {e}")
    except Exception as e:
        print(f"An unexpected error occurred while populating the payments sheet: {e}")

def populate_payments_sheet_2(payments_sheet_2, data_sheet):
    """
    Populate the payments sheet (sh_n=2) with data extracted from the data sheet.
    
    This function performs several tasks:
    1. Converts the source data into a pandas DataFrame.
    2. Validates the presence of required columns in the DataFrame.
    3. Aggregates the data.
    4. Extracts lockdown periods and generates dynamic column mappings.
    5. Updates sheet headings with actual lockdown periods.
    6. Inserts new rows and copies the formatting.
    7. Populates the target sheet with the aggregated data.
    8. Adds SUM formulas to calculate totals in the sheet.
    9. Restores merged cells and adjusts row heights.
    10. Applies conditional formatting to the new rows.
    11. Adjusts column visibility for certain ranges.
    
    Parameters:
    payments_sheet_2 (obj): The target sheet where data will be populated.
    data_sheet (obj): The source sheet from which data will be extracted.
    
    Returns:
    int: Number of rows added to the sheet.
    """
    try:
        # 1. Convert the source data to a pandas DataFrame
        data = convert_to_dataframe(data_sheet)

        # 2. Validate the presence of required columns in the DataFrame
        required_columns = ["IDNUMBER", "FIRSTNAME", "LASTNAME", "TERMINATIONDATE", "BANK_PAY_AMOUNT", "SHUTDOWN_TILL"]
        validate_columns(data, required_columns)

        # 3. Process and aggregate the data
        aggregated = aggregate_data_3_2(data)

        # 4. Extract lockdown periods and generate dynamic column mappings
        from tp_3_2 import extract_lockdown_periods_for_headings, generate_dynamic_month_columns, update_sheet_headings
        
        print("DEBUG: Extracting lockdown periods for dynamic headings...")
        period_headings = extract_lockdown_periods_for_headings(data)
        
        print("DEBUG: Generating dynamic month column mappings...")
        month_columns = generate_dynamic_month_columns(period_headings)
        
        # 5. Update sheet headings with actual lockdown periods
        print("DEBUG: Updating sheet headings...")
        update_sheet_headings(payments_sheet_2, period_headings)

        # 6. Prepare for row insertion
        num_rows_to_add = len(aggregated)
        merged_cells_to_restore = unmerge_cells_in_range(payments_sheet_2, start_row=18, end_row=31)

        # 7. Insert new rows into the target sheet
        insert_rows(payments_sheet_2, num_rows_to_add, insert_start_row=15)
        start_row = 15

        # 8. Copy formatting from row 14
        copy_formatting(payments_sheet_2, start_row, num_rows_to_add, source_cell_n=14)

        # 9. Populate the sheet with the aggregated data
        populate_sheet_3_2(payments_sheet_2, aggregated, month_columns)

        # 10. Add SUM formulas in columns G to V , Y to AO and AQ
        total_row = start_row + num_rows_to_add + 1  # 2nd row after the last inserted row
        for col in range(column_letter_to_index("G"), column_letter_to_index("V") + 1):
            col_letter = column_index_to_letter(col)
            sheet_range = f"{col_letter}{start_row}:{col_letter}{start_row + num_rows_to_add - 1}"
            payments_sheet_2[f"{col_letter}{total_row}"] = f"=SUM({sheet_range})"
        for col in range(column_letter_to_index("Y"), column_letter_to_index("AO") + 1):
            col_letter = column_index_to_letter(col)
            sheet_range = f"{col_letter}{start_row}:{col_letter}{start_row + num_rows_to_add - 1}"
            payments_sheet_2[f"{col_letter}{total_row}"] = f"=SUM({sheet_range})"
        for col in range(column_letter_to_index("AQ"), column_letter_to_index("AQ") + 1):
            col_letter = column_index_to_letter(col)
            sheet_range = f"{col_letter}{start_row}:{col_letter}{start_row + num_rows_to_add - 1}"
            payments_sheet_2[f"{col_letter}{total_row}"] = f"=SUM({sheet_range})"

        # Add formula in column W to sum G to V
        sum_g_to_v_range = f"G{total_row}:V{total_row}"
        payments_sheet_2[f"W{total_row}"] = f"=SUM({sum_g_to_v_range})"

        # Add formula in column AP to sum Y to AO
        sum_y_to_ao_range = f"Y{total_row}:AO{total_row}"
        payments_sheet_2[f"AP{total_row}"] = f"=SUM({sum_y_to_ao_range})"

        # 11. Update existing formulas to account for inserted rows (AFTER all operations)
        # REMOVED: update_formulas_after_row_insertion(payments_sheet_2, 15, num_rows_to_add)
        # Keeping original formulas as they are correct

        # 12. Restore any merged cells that were temporarily unmerged
        reapply_merged_cells(payments_sheet_2, merged_cells_to_restore, num_rows_to_add)

        # 13. Adjust row heights for better presentation
        reset_row_heights(
            payments_sheet_2,
            reference_row=14,
            target_rows=range(18, 32 + num_rows_to_add),
            hide_reference_row=True
        )

        # 14. Apply conditional formatting to new rows
        columns_to_format = [column_index_to_letter(i) for i in range(column_letter_to_index('A'), column_letter_to_index('AO') + 1)]
        legend_column = 'AS'
        apply_conditional_formatting_general(
            payments_sheet_2, 
            start_row, 
            num_rows_to_add, 
            columns_to_format, 
            legend=legend_column
        )

        # 15. Adjust column visibility for G-V and Y-AN ranges
        adjust_column_visibility(payments_sheet_2, start_row, start_row + num_rows_to_add - 1, 'G', 'V', 'Y', 'AN')
        
        return num_rows_to_add

    except KeyError as e:
        print(f"Error: Missing column during payments sheet population - {e}")
    except Exception as e:
        print(f"An unexpected error occurred while populating the payments sheet 2: {e}")

def populate_payments_sheet_3(payments_sheet_3, data_sheet):
    """
    Populate the Payments Sheet 3 with extracted data, format the columns, and reset rows as needed.
    
    This function performs several tasks:
    1. Converts the source data into a pandas DataFrame.
    2. Validates the presence of required columns in the data.
    3. Aggregates the data based on specific criteria.
    4. Inserts new rows into the target sheet.
    5. Copies formatting from a reference row to the newly inserted rows.
    6. Populates the sheet with the aggregated data.
    7. Applies conditional formatting to specific columns.
    8. Restores any merged cells that were temporarily unmerged.
    9. Adjusts the row heights for better presentation.
    
    Parameters:
    payments_sheet_3 (obj): The target sheet where data will be populated.
    data_sheet (obj): The source sheet from which data will be extracted.
    
    Raises:
    KeyError: If a required column is missing in the data.
    Exception: If an unexpected error occurs during processing.
    
    Returns:
    None
    """
    try:
        # 1. Convert the source data to a pandas DataFrame
        data = convert_to_dataframe(data_sheet)

        # 2. Validate the presence of required columns
        required_columns_3 = ["IDNUMBER", "FIRSTNAME", "LASTNAME"]
        validate_columns(data, required_columns_3)

        # 3. Aggregate the data based on specific criteria
        aggregated = aggregate_data_3_3(data)

        # 4. Calculate the number of rows to add
        num_rows_to_add = len(aggregated)
        merged_cells_to_restore = unmerge_cells_in_range(payments_sheet_3, start_row=13, end_row=23)

        # 5. Insert new rows into the target sheet at row 11
        start_row_3 = 11
        insert_rows(payments_sheet_3, num_rows_to_add, start_row_3)

        # 6. Copy formatting from row 10 to the newly inserted rows
        copy_formatting(payments_sheet_3, start_row_3, num_rows_to_add, source_cell_n=10)

        # 7. Populate the sheet with data using mappings
        populate_sheet_3_3(
            payments_sheet_3,
            aggregated,
            mappings={
                "A": lambda i, row: i,  # Row numbering
                "B": lambda i, row: row["IDNUMBER"],
                "C": lambda i, row: row["FIRSTNAME"],
                "D": lambda i, row: row["LASTNAME"]
            }
        )

        # 8. Format columns F and H with the general formatter
        columns_to_format = ['F', 'H']
        apply_conditional_formatting_general(payments_sheet_3, start_row_3, num_rows_to_add, columns_to_format, legend='K')

        # 9. Update existing formulas to account for inserted rows (AFTER all operations)
        # REMOVED: update_formulas_after_row_insertion(payments_sheet_3, 11, num_rows_to_add)
        # Keeping original formulas as they are correct

        # 10. Restore any merged cells that were temporarily unmerged
        reapply_merged_cells(payments_sheet_3, merged_cells_to_restore, num_rows_to_add)

        # 11. Reset row heights for rows 13 to 23
        reset_row_heights(
            payments_sheet_3, 
            reference_row=10, 
            target_rows=range(13, 24 + num_rows_to_add), 
            hide_reference_row=True
        )

    except KeyError as e:
        print(f"Error: Missing required column in data file - {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
