import streamlit as st
import pandas as pd

st.title("📊 Extract Table from B22 Until End Conditions")

uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

def extract_table(file):
    # Read from row 21 (B22), no header
    df = pd.read_excel(file, header=None, skiprows=21)
    
    # Start from column B (index 1) and drop completely empty columns
    df = df.iloc[:, 1:].dropna(axis=1, how='all')
    
    # Initialize variables
    stop_phrase = "ticket quantities will be rounded up in minimums and multiples of 100 pcs."
    cut_index = None
    
    # Find the first row where any stop condition is met
    for i in range(len(df)):
        row = df.iloc[i]
        
        # Condition 1: Check for stop phrase in any column
        row_text = " ".join(str(x).lower() for x in row if pd.notna(x))
        if stop_phrase in row_text:
            cut_index = i
            break
            
        # Condition 2: Check for "None" in first column (Style column)
        if len(row) > 0 and str(row.iloc[0]).strip().lower() == "none":
            cut_index = i
            break
            
        # Condition 3: Check for "total" in the row
        if "total" in row_text:
            cut_index = i
            break
    
    # If stop condition found, keep only rows before it
    if cut_index is not None:
        df = df.iloc[:cut_index]
    
    # Drop completely empty rows
    df = df.dropna(how='all')
    
    # Reset index to clean up the dataframe
    df = df.reset_index(drop=True)
    
    return df

# Apply for all uploaded files
if uploaded_files:
    for file in uploaded_files:
        table = extract_table(file)
        if not table.empty:
            st.subheader(f"📁 File: {file.name}")
            st.dataframe(table)
        else:
            st.warning(f"No valid table extracted from {file.name}")