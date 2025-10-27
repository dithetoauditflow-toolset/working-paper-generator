#tp_2_2.py
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule

START_ROW = 14

def aggregate_data_2_2(data):
    """
    Aggregates employee data for Sheet 2.

    This function groups the input data by "IDNUMBER" and calculates aggregate values for the employee's
    last name, first name, and employment start date. The result is then sorted by the last name.

    Args:
        data (pandas.DataFrame): The input data containing employee records.

    Returns:
        pandas.DataFrame: The aggregated data, sorted by last name.
    """
    # Aggregate the data
    aggregated_data = data.groupby("IDNUMBER").agg({
        "LASTNAME": "first",
        "FIRSTNAME": "first",
        "EMPLOYMENTSTARTDATE": "first"
    }).reset_index()
    
    # Sort the rows by LASTNAME
    aggregated_data = aggregated_data.sort_values(by="LASTNAME")
    
    return aggregated_data

def populate_sheet_2_2(employee_sheet, aggregated, mappings):
    """
    Populates the employee sheet with custom mappings.

    This function uses custom mappings to populate specific columns in the employee sheet. 
    It iterates over the aggregated data and applies the respective mapping functions to the 
    corresponding columns for each row.

    Args:
        employee_sheet (openpyxl.worksheet.worksheet.Worksheet): The worksheet object to populate with data.
        aggregated (pandas.DataFrame): The aggregated data to populate into the sheet.
        mappings (dict): A dictionary where keys are column letters and values are functions 
                         that define how to populate the corresponding columns.

    Returns:
        None
    """
    current_row = START_ROW
    for i, row in enumerate(aggregated.itertuples(index=False), start=1):
        try:
            for col_letter, func in mappings.items():
                employee_sheet[f"{col_letter}{current_row}"].value = func(i, row._asdict())
            current_row += 1
        except Exception as e:
            print(f"Error while populating custom mapped data: {e}")

def apply_conditional_formatting_2_2(sheet, start_row, end_row, columns, condition="No", fill_color="FFCCCC"):
    """
    Adds conditional formatting to specified columns based on the condition.

    This function applies conditional formatting to the specified columns in the worksheet. 
    It highlights cells with a specific value (e.g., "No") with a fill color.

    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): The sheet object where conditional formatting will be applied.
        start_row (int): The starting row for the formatting.
        end_row (int): The ending row for the formatting.
        columns (list): A list of column letters to apply the conditional formatting to.
        condition (str): The text value to trigger the conditional formatting (default: "No").
        fill_color (str): The hex color code for the fill color (default: "FFCCCC", light red).

    Returns:
        None
    """
    # Define red fill for "No"
    red_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

    # Apply to specified columns for "YES/NO"
    for col in columns:
        sheet.conditional_formatting.add(
            f"{col}{start_row}:{col}{end_row}",
            CellIsRule(operator="equal", formula=[f'"{condition}"'], fill=red_fill)
        )
