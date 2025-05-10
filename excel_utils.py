import pandas as pd

def get_sheet_names(filepath):
    # Returns -> A list of all sheet names.
    xls = pd.ExcelFile(filepath)
    return xls.sheet_names


def get_row_names(filepath, sheet_name):
    # Returns -> A list of row names from Excel file.
    # Raises ValueError -> If the sheet name is not found in the Excel file.
    xls = pd.ExcelFile(filepath)
    
    if sheet_name not in xls.sheet_names:
        raise ValueError(f"Sheet '{sheet_name}' not found.")
    
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    row_names = df.iloc[:, 0].dropna().astype(str).tolist()
    return row_names


def get_row_sum(filepath, sheet_name, row_name):
    # Returns -> The sum of all numeric values in a row, identified by its first column value.
    xls = pd.ExcelFile(filepath)
    
    if sheet_name not in xls.sheet_names:
        raise ValueError(f"Sheet '{sheet_name}' not found.")
    
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    # Find the row index where the first column matches the given row_name
    row_index = df[df.iloc[:, 0].astype(str).str.strip() == row_name.strip()].index

    if row_index.empty:
        raise ValueError(f"Row '{row_name}' not found.")
    
    row = df.iloc[row_index[0], 1:]  # Skip the first column (row name)
    numeric_values = pd.to_numeric(row, errors='coerce')  # Convert to numeric, ignoring non-numeric values
    return numeric_values.sum(skipna=True)  # Sum all numeric values, skipping NaN

