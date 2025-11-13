from turtle import st
import streamlit as st
import pandas as pd
import re
from io import BytesIO

def read_excel_table(excel_file):
    """Read tables from all sheets of an Excel file starting from A22, with specific stopping conditions"""
    try:
        all_sheets_data = []
        
        if excel_file.name.endswith(".xls"):
            xl_file = pd.ExcelFile(excel_file, engine='xlrd')
            sheet_names = xl_file.sheet_names
        else:
            xl_file = pd.ExcelFile(excel_file)
            sheet_names = xl_file.sheet_names
        
        for sheet_name in sheet_names:
            # Initialize variables
            qty_col_idx = None
            skip_rows = [22]  
            
            # Read the entire sheet first to check for stopping conditions
            df_full = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            
            # Check if header row (row 22) contains the stopping text
            header_has_stopping_text = False
            for cell in df_full.iloc[21]:  # Row 22 (0-indexed)
                if isinstance(cell, str) and "Ticket quantities will be rounded up in minimums and multiples of 100 pcs." in cell:
                    header_has_stopping_text = True
                    break
            
            if header_has_stopping_text:
                # Skip the entire sheet if header contains stopping text
                all_sheets_data.append({
                    'sheet_name': sheet_name, 
                    'data': pd.DataFrame(),  # Empty DataFrame
                    'style_number': None,
                    'stop_row': 21,  # Row 22
                    'qty_col_idx': None
                })
                continue
            
            # Find all rows (starting from row 22) that contain the stopping text
            stopping_rows = []
            for idx in range(21, len(df_full)):  # Start from row 22 (0-indexed)
                for cell in df_full.iloc[idx]:
                    if isinstance(cell, str) and "Ticket quantities will be rounded up in minimums and multiples of 100 pcs." in cell:
                        stopping_rows.append(idx)
                        break
            
            # Add all stopping rows to skip list
            skip_rows.extend(stopping_rows)
            
            # Find the QTY column in the header row (row 22)
            if len(df_full) > 21:  # Make sure we have at least 22 rows
                header_row = df_full.iloc[21]  # Row 22 (0-indexed)
                for idx, cell in enumerate(header_row):
                    if isinstance(cell, str) and "QTY" in cell.upper():
                        qty_col_idx = idx
                        break
            
            # Read the data with header at row 22, skipping specified rows
            usecols = range(qty_col_idx + 1) if qty_col_idx is not None else None
            df = pd.read_excel(
                excel_file, 
                sheet_name=sheet_name, 
                header=21,  # Row 22 is header
                skiprows=skip_rows,  # Skip row 23 and any row with stopping text
                usecols=usecols
            )
            
            # Remove unnamed columns (columns with "Unnamed" in the header)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            
            # Clean the data
            df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
            
            # Find the first blank row in the STYLE column and truncate
            if 'STYLE' in df.columns:
                # Find the first index where STYLE is blank or NaN
                blank_style_idx = None
                for idx, style_val in enumerate(df['STYLE']):
                    if pd.isna(style_val) or str(style_val).strip() == '':
                        blank_style_idx = idx
                        break
                
                # Truncate the DataFrame at the first blank STYLE
                if blank_style_idx is not None:
                    df = df.iloc[:blank_style_idx].reset_index(drop=True)
            
            # Extract style number
            style_number = None
            if not df.empty and 'STYLE' in df.columns:
                style_values = df['STYLE'].dropna().values
                if len(style_values) > 0:
                    style_number = str(style_values[0]).strip()
            
            # Determine the first stopping row for display
            first_stop_row = min(stopping_rows) if stopping_rows else None
            
            # Always add the sheet info with all required keys
            all_sheets_data.append({
                'sheet_name': sheet_name, 
                'data': df, 
                'style_number': style_number,
                'stop_row': first_stop_row,
                'qty_col_idx': qty_col_idx
            })
        
        return all_sheets_data
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return []
    

print(f"Type of st: {type(st)}")
print(f"st module: {st}")

def process_excel_table_data(all_table_data):
    """Process Excel table data from multiple files into a single table format"""
    try:
        all_excel_items = []
        
        for table_info in all_table_data:
            excel_df = table_info['data']
            excel_sheet = table_info['sheet_name']
            file_name = table_info.get('file_name', 'Unknown File')
            
            if excel_df.empty:
                continue
            
            excel_columns = list(excel_df.columns)
            column_mapping = {}
            
            for col in excel_columns:
                col_lower = str(col).lower()
                if 'style' in col_lower:
                    column_mapping['Style'] = col
                elif 'cc' in col_lower or 'color' in col_lower or 'colour' in col_lower:
                    column_mapping['Colour Code'] = col
                elif 'size' in col_lower:
                    column_mapping['Size'] = col
                elif 'qty' in col_lower or 'quantity' in col_lower:
                    column_mapping['Quantity'] = col
                elif 'retail' in col_lower and 'us' in col_lower:
                    column_mapping['Retail US'] = col
                elif 'retail' in col_lower and 'ca' in col_lower:
                    column_mapping['Retail CA'] = col
                elif 'sku' in col_lower:
                    column_mapping['SKU'] = col
                elif 'article' in col_lower:
                    column_mapping['Article'] = col
            
            for _, row in excel_df.iterrows():
                excel_item = {}
                
                for std_col, excel_col in column_mapping.items():
                    if excel_col in excel_df.columns:
                        value = row[excel_col]
                        if pd.notna(value):
                            if 'size' in std_col.lower():
                                value = convert_excel_size_codes(value)
                            excel_item[std_col] = str(value).strip()
                        else:
                            excel_item[std_col] = ""
                    else:
                        excel_item[std_col] = ""
                
                # Add file and sheet information
                excel_item['Source File'] = file_name
                excel_item['Source Sheet'] = excel_sheet
                
                for col in excel_columns:
                    if col not in column_mapping.values():
                        value = row[col]
                        if pd.notna(value):
                            excel_item[f"Excel {col}"] = str(value).strip()
                        else:
                            excel_item[f"Excel {col}"] = ""
                
                all_excel_items.append(excel_item)
        
        excel_data_df = pd.DataFrame(all_excel_items)
        
        if not excel_data_df.empty:
            excel_data_df = excel_data_df.replace("", float("nan"))
            excel_data_df = excel_data_df.dropna(axis=1, how='all')
            excel_data_df = excel_data_df.fillna("")
            
            if 'Size' in excel_data_df.columns:
                size_order = {"XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5}
                excel_data_df["Size_Order"] = excel_data_df["Size"].map(
                    lambda x: size_order.get(str(x).strip().upper(), 99)
                )
                excel_data_df = excel_data_df.sort_values("Size_Order").drop("Size_Order", axis=1)
            
            excel_data_df = excel_data_df.reset_index(drop=True)
        
        return excel_data_df
        
    except Exception as e:
        st.error(f"Error processing Excel table data: {str(e)}")
        return pd.DataFrame()
    
def read_multiple_excel_tables(excel_files):
    """
    Read tables from all sheets of multiple Excel files starting from A22, with specific stopping conditions
    
    Args:
        excel_files: List of uploaded Excel files
        
    Returns:
        List of dictionaries containing sheet data from all files
    """
    all_files_data = []
    
    for excel_file in excel_files:
        file_data = read_excel_table(excel_file)
        
        # Add file name to each sheet's data
        for sheet_data in file_data:
            sheet_data['file_name'] = excel_file.name
        
        all_files_data.extend(file_data)
    
    return all_files_data

def convert_excel_size_codes(size_value):
    """Convert numeric size codes from Excel to text sizes"""
    if pd.isna(size_value):
        return ""
    
    size_str = str(size_value).strip()
    if size_str.isdigit():
        size_code = int(size_str)
        size_mapping = {
            33901: "XS", 33902: "S", 33903: "M", 
            33904: "L", 33905: "XL", 33906: "XXL"
        }
        return size_mapping.get(size_code, size_str)
    return size_str

def clean_decimal_values(value):
    """Remove decimal part from numeric values (e.g., 197575744481.0 -> 197575744481)"""
    if pd.isna(value):
        return value
    
    value_str = str(value).strip()
    
    # Check if it's a numeric value with .0 at the end
    if value_str.endswith('.0'):
        try:
            # Try to convert to float then to int to remove decimal
            num = float(value_str)
            if num.is_integer():
                return str(int(num))
        except ValueError:
            pass
    
    return value_str

def remove_leading_zeros(value):
    """Remove leading zeros from a string value"""
    if pd.isna(value):
        return value
    
    value_str = str(value).strip()
    # Remove leading zeros but keep at least one digit if all are zeros
    if value_str and value_str[0] == '0':
        # Check if all characters are zeros
        if all(c == '0' for c in value_str):
            return '0'  # Keep one zero if all are zeros
        # Remove leading zeros
        return value_str.lstrip('0') or '0'  # Ensure we don't return empty string
    
    return value_str

def clean_retail_value(value):
    """Clean retail value by removing dollar signs and trailing zeros after decimal."""
    if pd.isna(value):
        return value
    s = str(value).strip()
    # Remove dollar signs
    s = s.replace('$', '')
    try:
        num = float(s)
        # If it's an integer, return integer string, else return float string without trailing zeros
        if num.is_integer():
            return str(int(num))
        else:
            # Convert to string and remove trailing zeros and potential decimal point
            s_num = str(num)
            if '.' in s_num:
                s_num = s_num.rstrip('0').rstrip('.')
            return s_num
    except:
        return s