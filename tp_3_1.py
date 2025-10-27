#tp_3_1.py
from datetime import datetime
import pandas as pd

def aggregate_data_3_1(data):
    """
    Aggregate data for the payments sheet.
    Extract the Month, Payment Date, PAY_REF_ITR_1, and calculate the total BANK_PAY_AMOUNT.

    Args:
        data: A pandas DataFrame containing the source data.

    Returns:
        A DataFrame with aggregated data.
    """
    # If PAYMENTDATE is already in datetime format, skip parsing
    if not pd.api.types.is_datetime64_any_dtype(data['PAYMENTDATE']):
        # Define a list of possible date formats
        date_formats = [
            "%d-%b-%Y",              # Example: 28-May-2020
            "%d-%b-%Y %I:%M:%S %p",  # Example: 28-May-2020 03:11:10 PM
            "%Y-%m-%d %H:%M:%S",     # Example: 2020-05-28 15:11:10
            "%d/%m/%Y",              # Example: 28/05/2020
            "%m/%d/%Y",              # Example: 05/28/2020
            "%d-%m-%Y",              # Example: 28-05-2020
            "%Y/%m/%d",              # Example: 2020/05/28
            "%d %b %Y",              # Example: 28 May 2020
        ]

        def parse_date(date_str):
            """Try to parse a date string using a list of formats."""
            for fmt in date_formats:
                try:
                    return datetime.strptime(date_str, fmt)
                except (ValueError, TypeError):
                    continue
            return None  # Return None if no format matches

        # Apply the custom date parsing function
        data["PAYMENTDATE"] = data["PAYMENTDATE"].apply(parse_date)

        # Check if any dates could not be parsed
        if data["PAYMENTDATE"].isnull().any():
            print("Warning: Some dates could not be parsed and were set to None.")
    
    # Ensure PAYMENTDATE is in datetime format
    data["PAYMENTDATE"] = pd.to_datetime(data["PAYMENTDATE"], errors="coerce")

    # Strip time component so grouping is consistent
    data["PAYMENTDATE"] = data["PAYMENTDATE"].dt.date  # Keeps only date part

    # Convert PAYMENTDATE to Month-Year format
    data["Month"] = pd.to_datetime(data["PAYMENTDATE"]).dt.strftime('%B %Y')

    # Grouping by PAY_REF_ITR_1 with corrected date formatting
    aggregated = (
        data.groupby("PAY_REF_ITR_1")
        .agg(
            Month=("Month", "first"),  
            PaymentDate=("PAYMENTDATE", "first"),  # Uses date only (no time)
            TotalBankPayAmount=("BANK_PAY_AMOUNT", "sum"),
        )
        .reset_index()
    )

    return aggregated

def populate_sheet_3_1(sheet, aggregated_data):
    """
    Populate the target sheet with aggregated data.

    Args:
        sheet: The target sheet object.
        aggregated_data: A pandas DataFrame with aggregated data.
    """
    for idx, row in enumerate(aggregated_data.itertuples(index=False), start=20):
        sheet[f"A{idx}"] = row.Month
        sheet[f"B{idx}"] = row.PaymentDate
        sheet[f"C{idx}"] = row.PAY_REF_ITR_1
        sheet[f"D{idx}"] = row.TotalBankPayAmount  # Insert the total bank pay amount
