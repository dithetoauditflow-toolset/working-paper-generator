#tp_3_2.py
from datetime import datetime
from helper_funcs import (
    column_index_to_letter,
    column_letter_to_index
)
import pandas as pd
from openpyxl.utils import column_index_from_string, get_column_letter

def extract_lockdown_periods_for_headings(data):
    """
    Extract unique lockdown periods from the data and format them for use as column headings.
    
    Args:
        data (DataFrame): The input data containing SHUTDOWN_FROM and SHUTDOWN_TILL columns.
        
    Returns:
        list: A list of formatted period strings in chronological order.
    """
    print("=" * 50)
    print("DEBUG: Starting extraction of lockdown periods for headings...")
    print("=" * 50)
    print(f"DEBUG: Data columns: {list(data.columns)}")
    print(f"DEBUG: Processing {len(data)} rows of data...")
    
    # Check if required columns exist
    if "SHUTDOWN_FROM" not in data.columns:
        print("ERROR: SHUTDOWN_FROM column not found!")
        return []
    if "SHUTDOWN_TILL" not in data.columns:
        print("ERROR: SHUTDOWN_TILL column not found!")
        return []
    
    # Define possible date formats for parsing
    date_formats = [
        "%Y-%m-%d %H:%M:%S",  # Standard datetime format
        "%Y-%m-%d",           # ISO date format
        "%d/%m/%Y",           # European date format
        "%m/%d/%Y",           # US date format
        "%d-%b-%Y",           # Day-Month-Year with abbreviated month name
        "%d %B %Y",           # Day-Month-Year with full month name
    ]
    
    periods = set()  # Use a set to store unique periods (to avoid duplicates)
    
    print(f"DEBUG: Processing {len(data)} rows of data...")
    
    # Loop through each row and extract shutdown periods
    for idx, row in data.iterrows():
        shutdown_from = row.get("SHUTDOWN_FROM")
        shutdown_till = row.get("SHUTDOWN_TILL")
        
        print(f"DEBUG: Row {idx}: FROM={shutdown_from}, TILL={shutdown_till}")
        
        if pd.notna(shutdown_from) and pd.notna(shutdown_till):
            from_date, till_date = None, None
            
            # Try parsing the dates using the available formats
            for fmt in date_formats:
                try:
                    if not from_date:
                        from_date = datetime.strptime(str(shutdown_from), fmt)
                    if not till_date:
                        till_date = datetime.strptime(str(shutdown_till), fmt)
                except ValueError:
                    continue
            
            # If parsing was successful, add the period to the set
            if from_date and till_date:
                period_str = f"{from_date.strftime('%d %B %Y')} to {till_date.strftime('%d %B %Y')}"
                periods.add((from_date, till_date, period_str))
                print(f"DEBUG: Found period: {period_str}")
            else:
                print(f"DEBUG: Could not parse dates: {shutdown_from}, {shutdown_till}")
        else:
            print(f"DEBUG: Row {idx} has missing dates")
    
    # Sort the periods by the 'from_date'
    sorted_periods = sorted(periods, key=lambda x: x[0])
    
    # Extract just the formatted strings
    period_headings = [period[2] for period in sorted_periods]
    
    print(f"DEBUG: Extracted {len(period_headings)} unique periods:")
    for i, period in enumerate(period_headings):
        print(f"DEBUG: Period {i+1}: {period}")
    
    print("=" * 50)
    return period_headings

def generate_dynamic_month_columns(period_headings):
    """
    Generate dynamic column mappings based on actual lockdown periods.
    
    Args:
        period_headings (list): List of formatted period strings.
        
    Returns:
        dict: Dictionary mapping period strings to Excel column letters.
    """
    print("DEBUG: Generating dynamic month column mappings...")
    print(f"DEBUG: Input period_headings: {period_headings}")
    
    # Define the available column ranges for the two sections
    # First section: G to V (16 columns) - for amounts claimed
    first_section_columns = ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']
    
    # Second section: Y to AO (17 columns) - for amounts paid
    second_section_columns = ['Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO']
    
    month_columns = {}
    
    # Map periods to first section columns (amounts claimed)
    for i, period in enumerate(period_headings[:len(first_section_columns)]):
        month_columns[period] = first_section_columns[i]
        print(f"DEBUG: First section (Claimed) - Mapped '{period}' to column {first_section_columns[i]}")
    
    # Map periods to second section columns (amounts paid)
    for i, period in enumerate(period_headings[:len(second_section_columns)]):
        # Use (PAID) suffix to match the sheet headers exactly
        period_paid = f"{period} (PAID)"
        month_columns[period_paid] = second_section_columns[i]
        print(f"DEBUG: Second section (Paid) - Mapped '{period_paid}' to column {second_section_columns[i]}")
    
    print(f"DEBUG: Total mappings created: {len(month_columns)}")
    return month_columns

def update_sheet_headings(sheet, period_headings):
    """
    Update the sheet headings in row 13 with the actual lockdown periods.
    
    Args:
        sheet: The Excel worksheet to update.
        period_headings (list): List of formatted period strings.
    """
    print("DEBUG: Updating sheet headings in row 13...")
    print(f"DEBUG: Input period_headings: {period_headings}")
    
    # Define the available column ranges for the two sections
    first_section_columns = ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']
    second_section_columns = ['Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO']
    
    print(f"DEBUG: First section columns (G-V): {first_section_columns}")
    print(f"DEBUG: Second section columns (Y-AO): {second_section_columns}")
    
    # Update first section headings (amounts claimed)
    for i, period in enumerate(period_headings[:len(first_section_columns)]):
        col_letter = first_section_columns[i]
        sheet[f"{col_letter}13"] = period
        print(f"DEBUG: First section (Claimed) - Updated {col_letter}13 to '{period}'")
    
    # Update second section headings (amounts paid)
    for i, period in enumerate(period_headings[:len(second_section_columns)]):
        col_letter = second_section_columns[i]
        sheet[f"{col_letter}13"] = f"{period} (PAID)"
        print(f"DEBUG: Second section (Paid) - Updated {col_letter}13 to '{period} (PAID)'")
    
    print(f"DEBUG: Total headings updated: {len(period_headings[:len(first_section_columns)])} in first section, {len(period_headings[:len(second_section_columns)])} in second section")
    print("DEBUG: Sheet headings update completed.")

def aggregate_data_3_2(data):
    """
    Aggregates employee data for the payments sheet by ensuring unique employees based on IDNUMBER 
    and summing BANK_PAY_AMOUNT for the same IDNUMBER and Period (lockdown period).
    """
    print("DEBUG: Starting aggregate_data_3_2 function...")
    
    # Debugging step: Check if BANK_PAY_AMOUNT exists
    if "BANK_PAY_AMOUNT" not in data.columns:
        raise ValueError("Error: Missing 'BANK_PAY_AMOUNT' column in the input data.")

    # Ensure BANK_PAY_AMOUNT is numeric and handle invalid entries
    def parse_amount(amount):
        if isinstance(amount, list) and len(amount) > 0:
            return float(amount[0])  # Extract first value if it's a list
        elif isinstance(amount, (int, float)):
            return float(amount)  # Use numeric value as is
        elif isinstance(amount, str) and amount.strip():  # Handle string values
            try:
                return float(amount.replace(",", "").strip())
            except ValueError:
                return 0  # Default to 0 if parsing fails
        return 0  # Default to 0 for other cases

    # Convert and clean BANK_PAY_AMOUNT
    data["BANK_PAY_AMOUNT"] = data["BANK_PAY_AMOUNT"].apply(parse_amount)

    # Convert SHUTDOWN_FROM and SHUTDOWN_TILL to datetime and create period strings
    data["SHUTDOWN_FROM_DT"] = pd.to_datetime(data["SHUTDOWN_FROM"], dayfirst=True, errors="coerce")
    data["SHUTDOWN_TILL_DT"] = pd.to_datetime(data["SHUTDOWN_TILL"], dayfirst=True, errors="coerce")
    
    # Create period strings in the format "dd Month yyyy to dd Month yyyy"
    def create_period_string(row):
        if pd.notna(row["SHUTDOWN_FROM_DT"]) and pd.notna(row["SHUTDOWN_TILL_DT"]):
            return f"{row['SHUTDOWN_FROM_DT'].strftime('%d %B %Y')} to {row['SHUTDOWN_TILL_DT'].strftime('%d %B %Y')}"
        return None
    
    data["Period"] = data.apply(create_period_string, axis=1)
    data["Period_Order"] = data["SHUTDOWN_FROM_DT"]

    # Aggregate rows by IDNUMBER and Period (sum BANK_PAY_AMOUNT)
    # REMOVED TERMINATIONDATE from aggregation - will be set to "IN SERVICE" for all employees
    grouped_data = (
        data.groupby(["IDNUMBER", "Period"])
        .agg({"BANK_PAY_AMOUNT": "sum", "Period_Order": "first", "FIRSTNAME": "first", "LASTNAME": "first"})
        .reset_index()
    )

    # Get unique periods in chronological order
    unique_periods = grouped_data["Period_Order"].dropna().sort_values().unique()
    period_names = []
    for period_dt in unique_periods:
        # Find the corresponding period string for this date
        matching_rows = grouped_data[grouped_data["Period_Order"] == period_dt]
        if not matching_rows.empty:
            period_names.append(matching_rows.iloc[0]["Period"])

    # Initialize the aggregated DataFrame with unique employees
    unique_employees = grouped_data.drop_duplicates(subset=["IDNUMBER"], keep="first")
    # Create a new column for termination status - all employees are "IN SERVICE"
    aggregated_data = unique_employees[["IDNUMBER", "FIRSTNAME", "LASTNAME"]].reset_index(drop=True)
    aggregated_data["TERMINATION_STATUS"] = "IN SERVICE"  # Hardcoded value for all employees

    # Define the available column ranges for the two sections
    first_section_columns = ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']
    second_section_columns = ['Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO']
    
    # Add columns for periods with proper naming based on section
    # Only create first section columns since second section is for manual entry
    for i, period in enumerate(period_names):
        if i < len(first_section_columns):
            # First section - amounts claimed
            aggregated_data[period] = 0.0
            print(f"DEBUG: Added first section column (Claimed): {period}")
        if i >= len(first_section_columns):
            # If more periods than available columns, they will be ignored
            print(f"DEBUG: Ignoring period {period} - exceeds available columns")
    
    print(f"DEBUG: Total periods to process: {len(period_names)}")
    print(f"DEBUG: First section capacity: {len(first_section_columns)}")
    print(f"DEBUG: Second section capacity: {len(second_section_columns)}")
    print(f"DEBUG: Periods going to first section: {period_names[:len(first_section_columns)]}")
    print(f"DEBUG: Periods going to second section: {period_names[len(first_section_columns):len(first_section_columns) + len(second_section_columns)]}")

    # Populate the period columns with aggregated BANK_PAY_AMOUNT for each employee and period
    for _, row in grouped_data.iterrows():
        employee_id = row["IDNUMBER"]
        period = row["Period"]
        amount = row["BANK_PAY_AMOUNT"]

        if pd.notna(employee_id) and pd.notna(period):
            # Populate claimed amounts in first section only
            if period in aggregated_data.columns:
                aggregated_data.loc[aggregated_data["IDNUMBER"] == employee_id, period] += amount
            
            # Second section (amounts paid) is left blank for manual entry by users

    print(f"DEBUG: Final aggregated_data columns: {list(aggregated_data.columns)}")
    return aggregated_data

def populate_sheet_3_2(sheet, aggregated_data, month_columns):
    """
    Populates an Excel sheet with the aggregated data, including mapping BANK_PAY_AMOUNT values 
    to their respective period columns, and handling skipped periods.

    :param sheet: The target Excel sheet to populate.
    :param aggregated_data: A pandas DataFrame with aggregated employee data.
    :param month_columns: A dictionary mapping period names to their respective Excel column letters.
    """
    print("DEBUG: Starting populate_sheet_3_2 function...")
    print(f"DEBUG: Month columns mapping: {month_columns}")
    print(f"DEBUG: Aggregated data columns: {list(aggregated_data.columns)}")
    
    # Normalize the column names (remove any leading/trailing spaces)
    aggregated_data.columns = aggregated_data.columns.str.strip()

    # Ensure period columns are treated as numeric (coerce errors to NaN)
    period_names = list(month_columns.keys())
    for period in period_names:
        if period in aggregated_data.columns:
            aggregated_data[period] = pd.to_numeric(aggregated_data[period], errors='coerce')
    
    # Normalize period names in the month_columns dictionary (remove extra spaces)
    normalized_period_columns = {period.strip(): col for period, col in month_columns.items()}

    for idx, row in enumerate(aggregated_data.itertuples(index=False), start=15):
        # Fill in employee data
        sheet[f"A{idx}"] = idx - 14 # Row numbering
        sheet[f"B{idx}"] = row.IDNUMBER
        sheet[f"C{idx}"] = "" # Blank column
        sheet[f"D{idx}"] = row.FIRSTNAME
        sheet[f"E{idx}"] = row.LASTNAME
        sheet[f"F{idx}"] = row.TERMINATION_STATUS  # Changed to TERMINATION_STATUS (always "IN SERVICE")

        # Insert BANK_PAY_AMOUNT into the correct column based on the period
        for i, period in enumerate(period_names):
            # Check if this period exists in aggregated_data for claimed amounts (first section only)
            if period in aggregated_data.columns:
                period_value = aggregated_data.loc[idx - 15, period] # Accessing the correct period column

                # Check for None values (if no data, leave it blank)
                if pd.isna(period_value):
                    print(f"DEBUG: No claimed data for {period} in row ID {row.IDNUMBER}. Leaving cell blank.")
                else:
                    # Insert the claimed amount into the first section
                    column_letter = normalized_period_columns.get(period)
                    if column_letter:
                        sheet[f"{column_letter}{idx}"] = period_value
                        print(f"DEBUG: Inserted claimed amount {period_value} for {period} in column {column_letter}{idx}")
            
            # Second section (amounts paid) is left blank for manual entry by users

def adjust_column_visibility(sheet, start_row, end_row, columns_range_start, columns_range_end, corresponding_columns_range_start, corresponding_columns_range_end):
    """
    Adjusts the visibility of columns in an Excel sheet based on whether they contain data.
    If a column in the first range is hidden, its corresponding column in the second range is also hidden.

    :param sheet: The Excel worksheet where the columns are located.
    :param start_row: The starting row of the range to check for data.
    :param end_row: The ending row of the range to check for data.
    :param columns_range_start: The starting column letter for the first range (e.g., 'G').
    :param columns_range_end: The ending column letter for the first range (e.g., 'V').
    :param corresponding_columns_range_start: The starting column letter for the second range (e.g., 'Y').
    :param corresponding_columns_range_end: The ending column letter for the second range (e.g., 'AN').
    """
    for col in range(column_letter_to_index(columns_range_start), column_letter_to_index(columns_range_end) + 1):
        col_letter = column_index_to_letter(col)
        corresponding_col_letter = column_index_to_letter(
            col - column_letter_to_index(columns_range_start) + column_letter_to_index(corresponding_columns_range_start)
        )
        
        # Check if any cell in the column has data in the specified range
        has_data = False
        for row in range(start_row, end_row + 1):
            if sheet[f"{col_letter}{row}"].value is not None:
                has_data = True
                break
        
        # Adjust visibility of the column
        sheet.column_dimensions[col_letter].hidden = not has_data
        sheet.column_dimensions[corresponding_col_letter].hidden = not has_data
        
def replicate_hidden_columns(source_sheet, target_sheet, start_letter, end_letter):
    """
    Replicate the hidden state of columns from source sheet to target sheet.
    
    Args:
        source_sheet: The source worksheet to copy column visibility from
        target_sheet: The target worksheet to copy column visibility to
        start_letter: Starting column letter (e.g., 'G')
        end_letter: Ending column letter (e.g., 'V')
    """
    try:
        start_idx = column_index_from_string(start_letter)
        end_idx = column_index_from_string(end_letter)
        
        for col_idx in range(start_idx, end_idx + 1):
            col_letter = get_column_letter(col_idx)
            # Copy the hidden state from the source sheet to the target sheet
            if col_letter in source_sheet.column_dimensions:
                target_sheet.column_dimensions[col_letter].hidden = source_sheet.column_dimensions[col_letter].hidden
            else:
                print(f"DEBUG: Column {col_letter} not found in source sheet")
    except Exception as e:
        print(f"DEBUG: Error replicating hidden columns: {e}")
