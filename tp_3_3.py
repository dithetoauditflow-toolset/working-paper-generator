#tp_3_3.py
import pandas as pd
from openpyxl import Workbook

def aggregate_data_3_3(data):
    """
    Aggregate the payments data for Sheet 3.
    Retain only the 'IDNUMBER', 'FIRSTNAME', and 'LASTNAME' columns, dropping duplicates by 'IDNUMBER'.
    """
    # Retain only necessary columns and drop duplicates
    aggregated_data = data[['IDNUMBER', 'FIRSTNAME', 'LASTNAME']].drop_duplicates(subset='IDNUMBER').reset_index(drop=True)
    
    # Sort the rows by LASTNAME
    aggregated_data = aggregated_data.sort_values(by="LASTNAME")
    
    return aggregated_data

def populate_sheet_3_3(payments_sheet, aggregated, mappings):
    """
    Populate the sheet with custom mappings.
    Writes data from the aggregated DataFrame to the specified sheet starting at row 11.
    """
    current_row = 11
    for i, row in enumerate(aggregated.itertuples(index=False), start=1):
        try:
            row_dict = row._asdict()
            for col_letter, func in mappings.items():
                cell = payments_sheet[f"{col_letter}{current_row}"]
                
                # Check if this cell is part of a merged range
                from openpyxl.utils import range_boundaries
                for merged_range in payments_sheet.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        # If it's part of a merged range, get the master cell (top-left cell)
                        min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
                        master_cell = payments_sheet.cell(row=min_row, column=min_col)
                        master_cell.value = func(i, row_dict)
                        break
                else:
                    # If it's not part of a merged range, set the value directly
                    cell.value = func(i, row_dict)
            current_row += 1
        except Exception as e:
            print(f"Error while populating custom mapped data: {e}")
