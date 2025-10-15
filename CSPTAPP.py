import streamlit as st
import pdfplumber
import openpyxl
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO
from fuzzywuzzy import fuzz
from datetime import datetime
import os

# -------------------- UI Configuration --------------------
st.set_page_config(
    page_title="Bandix Data Entry Checking Tool – Price Tickets - Razz Solutions",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------- Custom CSS --------------------
st.markdown("""
<style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Global Styles */
    .main {
        font-family: 'Inter', sans-serif;
    }
    
    /* Header Styling */
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
    }
    
    .main-title {
        color: white;
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    
    .main-subtitle {
        color: rgba(255,255,255,0.9);
        font-size: 1.2rem;
        font-weight: 400;
        margin: 0;
    }
    
    /* Status Cards */
    .status-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        border-left: 4px solid #4CAF50;
        margin: 1rem 0;
    }
    
    .status-card.warning {
        border-left-color: #FF9800;
    }
    
    .status-card.error {
        border-left-color: #f44336;
    }
    
    /* Section Headers */
    .section-header {
        background: linear-gradient(90deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 1rem 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #007bff;
        margin: 1.5rem 0 1rem 0;
    }
    
    .section-title {
        color: #2c3e50;
        font-size: 1.3rem;
        font-weight: 600;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* Upload Area Styling */
    .upload-container {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        border: 2px dashed #007bff;
        border-radius: 15px;
        padding: 2rem;
        text-align: center;
        margin: 1rem 0;
        transition: all 0.3s ease;
    }
    
    .upload-container:hover {
        border-color: #0056b3;
        background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
    }
    
    /* Metrics Cards */
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        text-align: center;
        box-shadow: 0 4px 20px rgba(0,0,0,0.1);
        border-top: 4px solid #007bff;
        transition: transform 0.3s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #2c3e50;
        margin-bottom: 0.5rem;
    }
    
    .metric-label {
        color: #6c757d;
        font-weight: 500;
        text-transform: uppercase;
        font-size: 0.9rem;
        letter-spacing: 1px;
    }
    
    /* Success/Warning/Error Messages */
    .alert-success {
        background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
        border: 1px solid #c3e6cb;
        color: #155724;
        padding: 1rem 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
        font-weight: 500;
    }
    
    .alert-warning {
        background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%);
        border: 1px solid #ffeaa7;
        color: #856404;
        padding: 1rem 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
        font-weight: 500;
    }
    
    .alert-info {
        background: linear-gradient(135deg, #d1ecf1 0%, #bee5eb 100%);
        border: 1px solid #bee5eb;
        color: #0c5460;
        padding: 1rem 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
        font-weight: 500;
    }
    
    /* Table Styling */
    .dataframe {
        border: none !important;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1) !important;
        border-radius: 10px !important;
        overflow: hidden !important;
    }
    
    /* Sidebar Styling */
    .css-1d391kg {
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
    }
    
    .css-1d391kg .css-1v0mbdj {
        color: white;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 2rem;
        margin-top: 3rem;
        border-top: 2px solid #e9ecef;
        color: #6c757d;
        font-style: italic;
    }
    
    /* Loading Spinner */
    .loading-container {
        display: flex;
        justify-content: center;
        align-items: center;
        padding: 2rem;
    }
    
    /* Progress Steps */
    .progress-steps {
        display: flex;
        justify-content: center;
        margin: 2rem 0;
        gap: 1rem;
    }
    
    .step {
        padding: 0.5rem 1rem;
        border-radius: 20px;
        background: #e9ecef;
        color: #6c757d;
        font-weight: 500;
        font-size: 0.9rem;
    }
    
    .step.active {
        background: #007bff;
        color: white;
    }
    
    .step.completed {
        background: #28a745;
        color: white;
    }
    
    /* Mismatch Summary Card */
    .mismatch-summary {
        background: linear-gradient(135deg, #fff5f5 0%, #ffe0e0 100%);
        border: 1px solid #ffcdd2;
        color: #c62828;
        padding: 1rem 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
        font-weight: 500;
    }
    
    .mismatch-example {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        margin-top: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    /* PIN Input Styling */
    .pin-container {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 10px;
        padding: 1rem;
        margin-top: 1rem;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }
    
    .pin-title {
        color: black;  /* Changed from white to black */
        font-size: 1rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    
    .search-container {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 10px;
        padding: 1rem;
        margin-top: 1rem;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }
</style>
""", unsafe_allow_html=True)

# -------------------- Main Header --------------------
st.markdown("""
<div class="main-header">
    <h1 class="main-title">🚀 Brandix Data Entry Checking Tool – Price Tickets</h1>
    <p class="main-subtitle">Advanced PO vs WO Comparison Dashboard | Powered by Razz </p>
</div>
""", unsafe_allow_html=True)

# -------------------- Function to log to Text File --------------------
def log_to_text(username, product_code, references, match_status, po_number=""):
    """Log analysis results to a daily text file"""
    try:
        # Define the directory for text files
        log_dir = r"C:\Users\APP\Desktop\ITL\CSAPP_Logs"
        
        # Create the directory if it doesn't exist
        os.makedirs(log_dir, exist_ok=True)
        
        # Get current date for filename
        current_date = datetime.now().strftime("%Y-%m-%d")
        file_path = os.path.join(log_dir, f"{current_date}.txt")
        
        # Get current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Prepare the log entry
        log_entry = f"{timestamp},{username},{product_code},{references},{po_number},{match_status}\n"
        
        # Append to the file
        with open(file_path, "a", encoding="utf-8") as f:
            f.write(log_entry)
        
        return True, f"Report successfully logged to {file_path}"
    except Exception as e:
        return False, f"Error logging to text file: {e}"

# -------------------- Function to Read Log File and Convert to Excel --------------------
def read_log_file_and_convert_to_excel(date_str):
    """Read a log file by date and convert to Excel with separate date and time columns"""
    try:
        # Define the directory for text files
        log_dir = r"C:\Users\APP\Desktop\ITL\CSAPP_Logs"
        file_path = os.path.join(log_dir, f"{date_str}.txt")
        
        # Check if the file exists
        if not os.path.exists(file_path):
            return None, f"Log file for {date_str} not found"
        
        # Read the text file
        with open(file_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        
        # Parse the lines into a DataFrame
        data = []
        for line in lines:
            if line.strip():
                parts = line.strip().split(",")
                if len(parts) >= 6:
                    # Split timestamp into date and time
                    timestamp = parts[0]
                    if " " in timestamp:
                        date_part, time_part = timestamp.split(" ", 1)
                    else:
                        date_part = timestamp
                        time_part = ""
                    
                    data.append({
                        "Date": date_part,
                        "Time": time_part,
                        "User Name": parts[1],
                        "Product code": parts[2],
                        "References": parts[3],
                        "PO Number": parts[4],
                        "Status": parts[5]
                    })
        
        # Create DataFrame
        df = pd.DataFrame(data)
        
        # Convert to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Log Data')
        output.seek(0)
        
        return output, f"Successfully converted log file for {date_str}"
    except Exception as e:
        return None, f"Error converting log file: {e}"

# -------------------- Excel Table Extraction Functions --------------------
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
            skip_rows = [22]  # Always skip row 23 (index 22)
            
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

def process_excel_table_data(all_table_data):
    """Process Excel table data into a single table format"""
    try:
        all_excel_items = []
        
        for table_info in all_table_data:
            excel_df = table_info['data']
            excel_sheet = table_info['sheet_name']
            
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

# -------------------- User Selection --------------------
with st.sidebar:
    st.markdown("### ⭐ User Selection")
    selected_user = st.selectbox(
        "Select your name:",
        ["", "👧Tarini", "👧udari", "👧Shaini", "👧Priyangi", "👧Uvini","👦Vihanga","👧Nimesha"],
        help="You must select a user to access the application"
    )
    
    if not selected_user:
        st.warning("⚠️ Please select a user to continue")
    
    # -------------------- PIN Authentication --------------------
    st.markdown("""
    <div class="pin-container">
        <div class="pin-title">🔐 Admin Access</div>
    </div>
    """, unsafe_allow_html=True)
    
    pin_input = st.text_input(
        "Enter PIN to access logs:",
        type="password",
        help="Enter the 4-digit PIN to access log search functionality"
    )
    
    # -------------------- Date Search (only visible if PIN is correct) --------------------
    if pin_input == "4391":
        st.markdown("""
        <div class="search-container">
            <div class="pin-title">🔍 Search Logs by Date</div>
        </div>
        """, unsafe_allow_html=True)
        
        search_date = st.date_input(
            "Select date to search:",
            help="Select the date for which you want to retrieve logs"
        )
        
        if st.button("Search Logs", key="search_logs"):
            date_str = search_date.strftime("%Y-%m-%d")
            excel_file, message = read_log_file_and_convert_to_excel(date_str)
            
            if excel_file:
                st.success(message)
                st.download_button(
                    label=f"⬇️ Download Log for {date_str}",
                    data=excel_file,
                    file_name=f"CS_AI_Tool_Log_{date_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.error(message)
    elif pin_input:
        st.error("❌ Incorrect PIN. Please try again.")

# -------------------- Progress Steps --------------------
def show_progress_steps(current_step=1):
    return ""

# -------------------- Pattern --------------------
style_pattern = re.compile(r"^\d{8}$")

# -------------------- Style to PDF --------------------
def create_styles_pdf(styles: list) -> BytesIO:
    doc = fitz.open()
    page = doc.new_page()
    title = "Extracted Style Numbers:\n\n"
    content = title + "\n".join(styles) if styles else "No style numbers found."
    rect = fitz.Rect(50, 50, 550, 800)
    page.insert_textbox(rect, content, fontsize=12, fontname="helv", align=0)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# -------------------- Merge PDF --------------------
def merge_pdfs(original_pdf: BytesIO, styles_pdf: BytesIO) -> BytesIO:
    pdf_out = fitz.open()
    pdf_styles = fitz.open(stream=styles_pdf.read(), filetype="pdf")
    pdf_orig = fitz.open(stream=original_pdf.read(), filetype="pdf")
    
    pdf_out.insert_pdf(pdf_styles)
    pdf_out.insert_pdf(pdf_orig)
    
    output = BytesIO()
    pdf_out.save(output)
    output.seek(0)
    return output

# -------------------- Missing Function: Extract Style Numbers from PO First Page --------------------
def extract_style_numbers_from_po_first_page(pdf_file):
    """Extract style numbers from the first page of PO PDF"""
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            if len(pdf.pages) > 0:
                first_page_text = pdf.pages[0].extract_text() or ""
                # Look for "Extracted Style Numbers:" section
                extracted_section_match = re.search(r'Extracted Style Numbers:\s*(.+)', first_page_text, re.IGNORECASE)
                if extracted_section_match:
                    extracted_styles = re.findall(r'\b\d{8}\b', extracted_section_match.group(1))
                    return extracted_styles
                
                # Fallback: look for any 8-digit numbers
                style_numbers = re.findall(r'\b\d{8}\b', first_page_text)
                return style_numbers
        return []
    except Exception as e:
        st.error(f"Error extracting style numbers from PO: {e}")
        return []

# -------------------- NEW: Extract PO Number from PDF --------------------
def extract_po_number(pdf_file):
    """Extract PO Number from PO PDF"""
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            # Check first page for PO Number
            if len(pdf.pages) > 0:
                first_page_text = pdf.pages[0].extract_text() or ""
                
                # Look for "PO Number:" pattern
                po_number_match = re.search(r'PO Number:\s*(\d+)', first_page_text)
                if po_number_match:
                    return po_number_match.group(1)
                
                # Alternative patterns
                patterns = [
                    r'P\.O\.\s*Number:\s*(\d+)',
                    r'Purchase Order Number:\s*(\d+)',
                    r'PO\s*#:\s*(\d+)',
                    r'PO\s*No:\s*(\d+)',
                    r'PO\s*Number:\s*(\d+)'
                ]
                
                for pattern in patterns:
                    match = re.search(pattern, first_page_text, re.IGNORECASE)
                    if match:
                        return match.group(1)
                
                # Fallback: Look for any 7-8 digit number on the right side of the page
                words = first_page_text.split()
                for i, word in enumerate(words):
                    if re.match(r'^\d{7,8}$', word):
                        if i > len(words) / 2:
                            return word
                
                # If still not found, try all pages
                for page in pdf.pages:
                    page_text = page.extract_text() or ""
                    for pattern in patterns:
                        match = re.search(pattern, page_text, re.IGNORECASE)
                        if match:
                            return match.group(1)
                
                # Last resort: look for any 7-8 digit number in the entire document
                full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
                numbers = re.findall(r'\b\d{7,8}\b', full_text)
                if numbers:
                    return numbers[0]
        return ""
    except Exception as e:
        st.error(f"Error extracting PO number: {e}")
        return ""

# -------------------- Helper Functions --------------------
def clean_quantity(qty_str):
    """Convert strings like '1,148.0000' or '465.0000' into floats with 4 decimal places"""
    if not qty_str:
        return 0.0000
    qty_str = str(qty_str).strip().replace(",", "")
    try:
        return round(float(qty_str), 4)
    except ValueError:
        return 0.0000

def truncate_after_sri_lanka(addr: str) -> str:
    """
    Extract the address up to and including:
    - The first occurrence of "Sri Lanka" (if found)
    - The second occurrence of "India" (if found twice)
    - The first occurrence of "India" (if found once)
    - Otherwise, the entire address
    Case-insensitive matching.
    """
    if not addr:
        return addr.strip()
    addr_lower = addr.lower()
    sri_lanka_pos = addr_lower.find("sri lanka")
    if sri_lanka_pos != -1:
        return addr[:sri_lanka_pos + len("sri lanka")].strip()
    india_positions = []
    search_pos = 0
    while True:
        india_pos = addr_lower.find("india", search_pos)
        if india_pos == -1:
            break
        india_positions.append(india_pos)
        search_pos = india_pos + 1
    if len(india_positions) >= 2:
        second_india_pos = india_positions[1]
        return addr[:second_india_pos + len("india")].strip()
    if len(india_positions) == 1:
        first_india_pos = india_positions[0]
        return addr[:first_india_pos + len("india")].strip()
    return addr.strip()

def clean_size(size_str):
    """Extract only the primary size from strings like 'S | P' or 'M / M'"""
    if not size_str:
        return ""
    size_str = str(size_str).strip().upper()
    if "|" in size_str:
        return size_str.split("|")[0].strip()
    elif "/" in size_str:
        return size_str.split("/")[0].strip()
    return size_str

def extract_wo_fields(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    delivery = ""
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if "Deliver To:" in ln:
            if i > 0:
                prev_line = lines[i-1].strip()
                if "Customer Delivery Name" in prev_line:
                    delivery = truncate_after_sri_lanka(re.sub(r"Deliver To:\s*", "", ln).strip())
                else:
                    if (any(char.isdigit() for char in prev_line) or 
                        any(indicator in prev_line.lower() for indicator in 
                            ["street", "st", "road", "rd", "avenue", "ave", "building", "block", "no", "#"]) or
                        len(prev_line.split()) > 3):
                        combined = prev_line + " " + re.sub(r"Deliver To:\s*", "", ln).strip()
                        delivery = truncate_after_sri_lanka(combined)
                    else:
                        delivery = truncate_after_sri_lanka(re.sub(r"Deliver To:\s*", "", ln).strip())
            else:
                delivery = truncate_after_sri_lanka(re.sub(r"Deliver To:\s*", "", ln).strip())
            delivery = re.sub(r'PTK ENTERPRISES,\s*', '', delivery, flags=re.IGNORECASE)
            break
    codes = []
    for line in lines:
        if "Product Code" in line:
            found = re.findall(r"Product Code[:\s]*([\w\s\-]+(?:\s*/\s*[\w\s\-]+)*)", line)
            for match in found:
                for code in match.split("/"):
                    clean = code.strip().upper()
                    clean = clean.replace("VSBA", "")
                    clean = clean.replace("-", " ")
                    if clean:
                        codes.append(clean)
            break
    po_numbers = list(set(re.findall(r'\b\d{7,8}\b', text)))
    return {
        "customer_name": "",
        "delivery_address": delivery,
        "product_codes": list(set(codes)),
        "po_numbers": po_numbers
    }
    
def extract_po_fields(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    lines = [ln.strip() for ln in text.split("\n")]
    capture = False
    address_lines = []
    blank_count = 0
    for ln in lines:
        if "Delivery Location:" in ln:
            capture = True
            continue
        if capture:
            if "Forwarder:" in ln:
                break
            if not ln:
                blank_count += 1
                if blank_count >= 2:
                    break
                continue
            else:
                blank_count = 0
            address_lines.append(ln)
    if address_lines:
        full_address = " ".join(address_lines)
    else:
        plot_found = False
        temp_address_lines = []
        for i, line in enumerate(lines):
            if re.search(r'\b(plot|building)\s*#?\s*\d+', line, re.IGNORECASE):
                plot_found = True
                temp_address_lines = [line]
                for j in range(i + 1, len(lines)):
                    next_line = lines[j].strip()
                    if not next_line:
                        continue
                    temp_address_lines.append(next_line)
                    if re.search(r'\bindia\b', next_line, re.IGNORECASE):
                        break
                    if any(keyword in next_line.lower() for keyword in 
                           ['forwarder', 'shipping', 'invoice', 'payment', 'terms']):
                        temp_address_lines.pop()
                        break
                break
        if plot_found and temp_address_lines:
            full_address = " ".join(temp_address_lines)
        else:
            full_address = ""
    full_address = re.sub(r'\s+', ' ', full_address)
    full_address = full_address.replace('\xa0', ' ')
    full_address = full_address.replace('\u2013', '-')
    full_address = full_address.replace('\u2014', '-')
    final_addr = truncate_after_sri_lanka(full_address)
    final_addr = final_addr.strip()
    po_codes = re.findall(r"(LB\s*\d+)", text)
    sup_ref_codes = re.findall(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z]+[-\s]?\d+)", text, re.IGNORECASE)
    tag_codes = re.findall(r"TAG\.PRC\.TKT_(.*?)_REG", text)
    all_product_codes = list(set([c.strip().upper() for c in sup_ref_codes + tag_codes]))
    return {
        "delivery_location": final_addr,
        "product_codes": po_codes + all_product_codes,
        "all_found_addresses": [final_addr]
    }

def clean_wo_address(addr: str) -> str:
    if not addr:
        return addr.strip()
    addr = re.sub(r',\s*,\s*India\s*$', '', addr, flags=re.IGNORECASE)
    addr = re.sub(r',,\s*', ', ', addr)
    addr = re.sub(r'\s*,\s*', ', ', addr)
    addr = re.sub(r',\s*$', '', addr)
    return addr.strip()

def compare_addresses(wo, po):
    ns = fuzz.token_sort_ratio(wo["customer_name"], po["delivery_location"])
    as_ = fuzz.token_sort_ratio(wo["delivery_address"], po["delivery_location"])
    comb = max(ns, as_)
    return {
        "WO Name": wo["customer_name"], 
        "WO Addr": wo["delivery_address"], 
        "PO Addr": po["delivery_location"],
        "Name %": ns, 
        "Addr %": as_, 
        "Overall %": comb, 
        "Status": "✅ Match" if comb >= 90 else "⚠️ Review"
    }

def extract_style_numbers(po_pdf_path):
    style_numbers = set()
    with pdfplumber.open(po_pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        if row:
                            for cell in row:
                                if cell:
                                    cell_clean = str(cell).strip().replace(',', '')
                                    if "." in cell_clean:
                                        cell_clean = cell_clean.split(".")[0]
                                    if cell_clean.isdigit():
                                        val = int(cell_clean)
                                        if 10 < val < 100000:
                                            if val > qty:
                                                qty = val
        full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        found_free_text_styles = re.findall(r'\b[A-Z]{2,}\s*\d{3,}\b', full_text)
        style_numbers.update([s.strip() for s in found_free_text_styles])
        if len(pdf.pages) > 0:
            first_page_text = pdf.pages[0].extract_text() or ""
            extracted_section_match = re.search(r'Extracted Style Numbers:\s*(.+)', first_page_text)
            if extracted_section_match:
                extracted_styles = re.findall(r'\b[A-Z]{2,}\s*\d{3,}\b', extracted_section_match.group(1))
                style_numbers.update([s.strip() for s in extracted_styles])
    return list(style_numbers)

def reorder_wo_by_size(wo_items):
    size_order = {"XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5}
    def get_order(wo):
        size = wo.get("Size 1", "").strip().upper()
        return size_order.get(size, 99)
    return sorted(wo_items, key=get_order)

def reorder_po_by_size(po_details):
    size_order = {"XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5}
    def get_order(po):
        size = po.get("Size", "").strip().upper()
        return size_order.get(size, 99)
    return sorted(po_details, key=get_order)

def sort_items_by_size(items):
    size_order = {"XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5}
    def get_size_key(item):
        size = item.get("WO Size", item.get("PO Size", "")).strip().upper()
        return size_order.get(size, 99)
    return sorted(items, key=get_size_key)

# -------------------- MINIMAL, NON-BREAKING FIX inside this function ONLY --------------------
def extract_po_details(pdf_file):
    """Enhanced function to handle multiple PO formats with quantity aggregation"""
    pdf_file.seek(0)
    extracted_styles = extract_style_numbers_from_po_first_page(pdf_file)
    repeated_style = extracted_styles[0] if extracted_styles else ""
    pdf_file.seek(0)
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    has_tag_format = "TAG.PRC.TKT_" in text and "Color/Size/Destination :" in text
    has_original_format = any("Colour/Size/Destination:" in line for line in lines) or re.search(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z]+[-\s]?\d+)", text, re.IGNORECASE)
    po_items = []
    item_dict = {}  # Dictionary to aggregate quantities by size, color, and style
    if has_tag_format and not has_original_format:
        tag_match = re.search(r"TAG\.PRC\.TKT_(.*?)_REG", text)
        product_code_used = tag_match.group(1).strip().upper() if tag_match else ""
        product_code_used = product_code_used.replace("-", " ")
        i = 0
        while i < len(lines):
            line = lines[i]
            item_match = re.match(r'^(\d+)\s+TAG\.PRC\.TKT_.*?([\d,]+\.\d+)\s+PCS', line)
            if item_match:
                item_no = item_match.group(1)
                quantity_str = item_match.group(2)
                quantity = clean_quantity(quantity_str)
                colour = size = ""
                for j in range(i + 1, min(i + 5, len(lines))):
                    next_line = lines[j]
                    if "Color/Size/Destination :" in next_line:
                        cs_part = next_line.split(":", 1)[1].strip()
                        cs_parts = [part.strip() for part in cs_part.split(" / ") if part.strip()]
                        if len(cs_parts) >= 2:
                            size_keywords = ["XS", "S", "M", "L", "XL", "XXL", "XXXL", "XXG", "P", "G"]
                            first_part = cs_parts[0].strip()
                            if first_part.upper() in size_keywords:
                                size = first_part.upper()
                                colour = cs_parts[1].split()[0] if cs_parts[1] else ""
                            else:
                                colour_part = cs_parts[0].strip()
                                colour = colour_part.split()[0] if colour_part else ""
                                size = cs_parts[1].strip().upper()
                        
                        # NEW LOGIC: Check for size before last slash
                        # Find the last occurrence of a slash
                        last_slash_index = cs_part.rfind('/')
                        if last_slash_index != -1:
                            # Get the substring before the last slash
                            before_slash = cs_part[:last_slash_index].strip()
                            # Split into tokens and check in reverse order
                            tokens = before_slash.split()
                            for token in reversed(tokens):
                                if token.upper() in size_keywords:
                                    size = token.upper()
                                    break
                        break
                item_key = (size, colour.upper() if colour else "", repeated_style)
                if item_key in item_dict:
                    item_dict[item_key]["Quantity"] += quantity
                else:
                    item_dict[item_key] = {
                        "Item_Number": item_no,
                        "Item_Code": f"TAG_{product_code_used}",
                        "Quantity": quantity,
                        "Colour_Code": colour.upper() if colour else "",
                        "Size": size,
                        "Style 2": repeated_style,
                        "Product_Code": product_code_used,
                    }
            i += 1
    else:
        # ORIGINAL FORMAT HANDLING (kept original logic)
        sup_ref_match = re.search(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z]+[-\s]?\d+)", text, re.IGNORECASE)
        sup_ref_code = sup_ref_match.group(1).strip().upper() if sup_ref_match else ""
        sup_ref_code = sup_ref_code.replace("-", " ")
        tag_code = ""
        for i, line in enumerate(lines):
            if "Item Description" in line:
                if i + 2 < len(lines):
                    second_line = lines[i + 2]
                    match = re.search(r"TAG\.PRC\.TKT_(.*?)_REG", second_line)
                    if match:
                        tag_code = match.group(1).strip().upper()
                        tag_code = tag_code.replace("-", " ")
                break
        product_code_used = sup_ref_code if sup_ref_code else tag_code
        for i, line in enumerate(lines):
            # Strict pattern (original)
            item_match = re.match(r'^(\d+)\s+([A-Z0-9]+)\s+(\d+)\s+([\d,]+\.\d+)\s+PCS', line)
            # Fallback pattern (only used if strict fails) — non-breaking
            relaxed_match = None
            if not item_match:
                relaxed_match = re.match(r'^(\d+)\s+([A-Z0-9]+)\b.*?([\d,]+(?:\.\d+)?)\s+PCS', line)
            if item_match or relaxed_match:
                if item_match:
                    item_no, item_code, _, qty_str = item_match.groups()
                else:
                    item_no, item_code, qty_str = relaxed_match.groups()
                quantity = clean_quantity(qty_str)
                colour = size = ""
                # Keep original logic, just allow "Color/Size/Destination :" as additional fall-back
                for j in range(i + 1, min(i + 10, len(lines))):
                    ln = lines[j]
                    if not colour and ("Colour/Size/Destination:" in ln or "Color/Size/Destination :" in ln):
                        cs = ln.split(":", 1)[1].strip()
                        size_keywords = ["XS", "S", "M", "L", "XL", "XXL", "XXXL", "XXG", "P", "G"]
                        parts = [p.strip() for p in cs.split("/") if p.strip()]
                        if parts:
                            size_part = parts[0].split("|")[0].strip().upper()
                            if size_part in size_keywords:
                                size = size_part
                                if len(parts) > 1:
                                    colour = parts[1].strip().split()[0].strip().upper()
                            else:
                                colour = re.match(r'^(\S+)', cs).group(1) if re.match(r'^(\S+)', cs) else ""
                                size_match = re.search(r'/\s*([^/]+)\s*/', cs)
                                if size_match:
                                    size = size_match.group(1).strip()
                        
                        # NEW LOGIC: Check for size before last slash
                        # Find the last occurrence of a slash
                        last_slash_index = cs.rfind('/')
                        if last_slash_index != -1:
                            # Get the substring before the last slash
                            before_slash = cs[:last_slash_index].strip()
                            # Split into tokens and check in reverse order
                            tokens = before_slash.split()
                            for token in reversed(tokens):
                                if token.upper() in size_keywords:
                                    size = token.upper()
                                    break
                item_key = (size.upper() if size else "", (colour or "").strip().upper(), repeated_style)
                if item_key in item_dict:
                    item_dict[item_key]["Quantity"] += quantity
                else:
                    item_dict[item_key] = {
                        "Item_Number": item_no,
                        "Item_Code": item_code,
                        "Quantity": quantity,
                        "Colour_Code": (colour or "").strip().upper(),
                        "Size": (size or "").strip().upper(),
                        "Style 2": repeated_style,
                        "Product_Code": product_code_used,
                    }
    po_items = list(item_dict.values())
    return po_items

# -------------------- UPDATED WO Extraction Function --------------------
def extract_wo_items_table(pdf_file, product_codes=None):
    """
    Enhanced function to extract WO items from Victoria's Secret price ticket tables
    with improved table detection and data extraction for all formats, including sizes split across lines
    """
    items = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # First try standard table extraction
            tables = page.extract_tables()
            
            # If standard extraction fails, try with explicit lines
            if not tables or len(tables) == 0:
                tables = page.extract_tables({
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "explicit_vertical_lines": page.curves + page.edges,
                    "explicit_horizontal_lines": page.curves + page.edges,
                })
            
            # Process each table
            for table_idx, table in enumerate(tables):
                if not table or len(table) < 2:
                    continue
                
                # Pre-process table to handle sizes split across cells and within cells
                processed_table = []
                for row in table:
                    if not row:
                        continue
                    
                    processed_row = []
                    i = 0
                    while i < len(row):
                        cell = str(row[i]) if row[i] is not None else ""
                        
                        # Check if this cell ends with a slash and the next cell contains a size suffix
                        if i < len(row) - 1 and cell.strip().endswith("/"):
                            next_cell = str(row[i+1]) if row[i+1] is not None else ""
                            # Check if next cell is a size suffix (XP, P, M, G, XG)
                            if next_cell.strip().upper() in ["XP", "P", "M", "G", "XG"]:
                                # Combine the cells
                                combined_cell = cell + next_cell
                                processed_row.append(combined_cell)
                                i += 2  # Skip the next cell
                                continue
                        
                        # If not a split size, just add the cell as-is
                        processed_row.append(cell)
                        i += 1
                    
                    processed_table.append(processed_row)
                
                # Now find the header row
                header_row_idx = -1
                column_positions = {}
                
                for i, row in enumerate(processed_table):
                    if not row:
                        continue
                    
                    row_text = " ".join([str(cell).strip() for cell in row if cell])
                    if any(term in row_text for term in ["Style", "Colour Code", "Size", "Quantity"]):
                        header_row_idx = i
                        
                        # Map column positions
                        for j, cell in enumerate(row):
                            cell_text = str(cell).strip().lower() if cell else ""
                            if "style" in cell_text:
                                column_positions["style"] = j
                            elif "colour" in cell_text or "color" in cell_text:
                                column_positions["color_code"] = j
                            elif "size 1" in cell_text or "size" in cell_text:
                                column_positions["size1"] = j
                            elif "size 2" in cell_text:
                                column_positions["size2"] = j
                            elif "panty" in cell_text:
                                column_positions["panty_length"] = j
                            elif "retail" in cell_text and "us" in cell_text:
                                column_positions["retail_us"] = j
                            elif "retail" in cell_text and "ca" in cell_text:
                                column_positions["retail_ca"] = j
                            elif "multi" in cell_text:
                                column_positions["multi_price"] = j
                            elif "sku" in cell_text:
                                column_positions["sku"] = j
                            elif "article" in cell_text:
                                column_positions["article"] = j
                            elif "quantity" in cell_text or "qty" in cell_text:
                                column_positions["quantity"] = j
                        break
                
                # If we couldn't find a header row, try to infer it
                if header_row_idx == -1:
                    for i, row in enumerate(processed_table):
                        if not row or len(row) < 8:
                            continue
                        
                        first_cell = str(row[0]).strip()
                        if re.match(r'^\d{8}$', first_cell):
                            has_size = False
                            for cell in row:
                                cell_str = str(cell).strip().upper()
                                # Check for any size format, including combined ones
                                if any(size in cell_str for size in ["XS/XP", "S/P", "M/M", "L/G", "XL/XG", "XXL", "XXXL", "XXG", "XG", "XS", "S", "M", "L", "XL", "P", "G"]):
                                    has_size = True
                                    break
                            
                            if has_size:
                                header_row_idx = i
                                column_positions = {
                                    "style": 0,
                                    "color_code": 1,
                                    "size1": 2,
                                    "size2": 3,
                                    "panty_length": 4,
                                    "retail_us": 5,
                                    "retail_ca": 6,
                                    "multi_price": 7,
                                    "sku": 8,
                                    "article": 9,
                                    "quantity": len(row) - 1
                                }
                                break
                
                # Skip if we couldn't determine the header
                if header_row_idx == -1:
                    continue
                
                # Process data rows
                for row_idx, row in enumerate(processed_table[header_row_idx + 1:], header_row_idx + 1):
                    if not row or len(row) < max(column_positions.values()) + 1:
                        continue
                    
                    try:
                        style = str(row[column_positions.get("style", 0)] or "").strip()
                        color_code = str(row[column_positions.get("color_code", 1)] or "").strip().upper()
                        
                        # Handle size extraction with special care for combined sizes and multi-line sizes
                        size1_raw = str(row[column_positions.get("size1", 2)] or "").strip()
                        
                        # Check if the size cell contains a newline (like "XS\nXP")
                        if "\n" in size1_raw:
                            # Split by newline and take the first part
                            size_parts = size1_raw.split("\n")
                            # Process each part to handle the case where one part is just "/" and the next is "XP"
                            processed_size = ""
                            for part in size_parts:
                                if part.strip() == "/":
                                    processed_size += "/"
                                else:
                                    processed_size += part.strip()
                            
                            # Now clean the processed size
                            size1 = clean_size(processed_size)
                        else:
                            size1 = clean_size(size1_raw)
                        
                        # If we didn't get a valid size, try to find it in other cells
                        if not size1:
                            # Check each cell for size patterns
                            for cell in row:
                                cell_str = str(cell).strip()
                                
                                # Check if the cell contains a newline
                                if "\n" in cell_str:
                                    # Split by newline and check each part
                                    parts = cell_str.split("\n")
                                    processed_cell = ""
                                    for part in parts:
                                        if part.strip() == "/":
                                            processed_cell += "/"
                                        else:
                                            processed_cell += part.strip()
                                    
                                    # Look for size patterns in the processed cell
                                    cell_upper = processed_cell.upper()
                                    size_match = re.search(r'\b(XS/XP|S/P|M/M|L/G|XL/XG|XXL|XXXL|XXG|XG|XS|S|M|L|XL|P|G)\b', cell_upper)
                                    if size_match:
                                        size1 = clean_size(size_match.group(1))
                                        break
                                else:
                                    # Look for size patterns in the whole cell
                                    cell_upper = cell_str.upper()
                                    size_match = re.search(r'\b(XS/XP|S/P|M/M|L/G|XL/XG|XXL|XXXL|XXG|XG|XS|S|M|L|XL|P|G)\b', cell_upper)
                                    if size_match:
                                        size1 = clean_size(size_match.group(1))
                                        break
                        
                        # Extract quantity
                        quantity_str = ""
                        if "quantity" in column_positions:
                            quantity_str = str(row[column_positions["quantity"]] or "").strip()
                        
                        if not quantity_str or not re.search(r'\d', quantity_str):
                            # Try the last column as a fallback
                            quantity_str = str(row[-1] or "").strip()
                        
                        quantity = clean_quantity(quantity_str)
                        
                        # Only add item if we have valid style and quantity
                        if style and quantity > 0:
                            item_data = {
                                "Style": style,
                                "WO Colour Code": color_code,
                                "Size 1": size1,
                                "Quantity": quantity,
                                "WO Product Code": " / ".join(product_codes) if product_codes else ""
                            }
                            
                            # Add optional fields if they exist
                            if "size2" in column_positions:
                                size2_raw = str(row[column_positions["size2"]] or "").strip()
                                # Check if the size2 cell contains a newline
                                if "\n" in size2_raw:
                                    # Split by newline and process each part
                                    size_parts = size2_raw.split("\n")
                                    processed_size = ""
                                    for part in size_parts:
                                        if part.strip() == "/":
                                            processed_size += "/"
                                        else:
                                            processed_size += part.strip()
                                    
                                    item_data["Size 2"] = clean_size(processed_size)
                                else:
                                    item_data["Size 2"] = clean_size(size2_raw)
                            
                            if "panty_length" in column_positions:
                                item_data["Panty Length"] = str(row[column_positions["panty_length"]] or "").strip()
                            
                            if "retail_us" in column_positions:
                                item_data["Retail US"] = str(row[column_positions["retail_us"]] or "").strip()
                            
                            if "retail_ca" in column_positions:
                                item_data["Retail CA"] = str(row[column_positions["retail_ca"]] or "").strip()
                            
                            if "multi_price" in column_positions:
                                item_data["Multi Price"] = str(row[column_positions["multi_price"]] or "").strip()
                            
                            if "sku" in column_positions:
                                item_data["SKU"] = str(row[column_positions["sku"]] or "").strip()
                            
                            if "article" in column_positions:
                                item_data["Article"] = str(row[column_positions["article"]] or "").strip()
                            
                            items.append(item_data)
                    
                    except (ValueError, IndexError):
                        continue
    
    # If we still don't have items, try text-based extraction
    if not items:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
        
        # Try to find patterns in the text
        # Look for lines that contain style, color, size, and quantity
        lines = full_text.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Pattern: 8-digit style, then alphanumeric code, then size, then numbers
            # Example: 11276771 91S3 XS/ XP $14.50 $14.50 0198765613549 27184055 1605
            # We need to handle the space in "XS/ XP" and also newlines
            
            # First, normalize the line by removing extra spaces and handling newlines
            normalized_line = re.sub(r'\s+', ' ', line)
            
            # Try to match the pattern
            pattern = r'(\d{8})\s+([A-Z0-9]+)\s+(XS/XP|S/P|M/M|L/G|XL/XG|XXL|XXXL|XXG|XG|XS|S|M|L|XL|P|G)\s+\$[\d.]+\s+\$[\d.]+\s+\d+\s+\d+\s+(\d{1,4}(?:,\d{3})*)'
            match = re.search(pattern, normalized_line)
            
            if match:
                style = match.group(1)
                color_code = match.group(2)
                size = match.group(3)
                quantity = match.group(4).replace(',', '')
                
                try:
                    items.append({
                        "Style": style,
                        "WO Colour Code": color_code.upper(),
                        "Size 1": clean_size(size),
                        "Quantity": int(quantity),
                        "WO Product Code": " / ".join(product_codes) if product_codes else ""
                    })
                except ValueError:
                    continue
    
    # Aggregate quantities for items with same style, color code, and size
    aggregated_items = {}
    for item in items:
        key = (item["Style"], item["WO Colour Code"], item["Size 1"])
        if key in aggregated_items:
            aggregated_items[key]["Quantity"] += item["Quantity"]
        else:
            aggregated_items[key] = item.copy()
    
    # Convert back to list
    final_items = list(aggregated_items.values())
    
    return final_items

def _extract_qty_from_text(text, min_value=1, max_value=100000):
    if not text:
        return None
    candidates = re.findall(r'\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?', str(text))
    best = None
    for token in candidates:
        num_str = token.replace(",", "")
        try:
            int_part = num_str.split(".", 1)[0]
            if not int_part:
                continue
            val = int(int_part)
            if min_value <= val <= max_value:
                if best is None or val > best:
                    best = val
        except ValueError:
            continue
    return best

def _dedupe(items):
    seen = set()
    unique = []
    for it in items:
        key = (it["Style"], it.get("WO Colour Code", ""), it.get("Size 1", ""), it["Quantity"])
        if key in seen:
            continue
        seen.add(key)
        unique.append(it)
    return unique

def extract_wo_items_table_enhanced(pdf_file, product_codes=None):
    return extract_wo_items_table(pdf_file, product_codes)

def enhanced_quantity_matching(wo_items, po_details, tolerance=0):
    matched, mismatched = [], []
    used = set()
    for wo in wo_items:
        wq = wo["Quantity"]
        ws = wo.get("Size 1", "").strip().upper()
        wc = wo.get("WO Colour Code", "").strip().upper()
        wstyle = wo.get("Style", "").strip()
        full_match_found = False
        partial_match_idx = None
        partial_match_score = -1
        for idx, po in enumerate(po_details):
            if idx in used:
                continue
            pq = po["Quantity"]
            ps = po.get("Size", "").strip().upper()
            pc = po.get("Colour_Code", "").strip().upper()
            pstyle = po.get("Style 2", "").strip()
            if wq == pq and ws == ps and wc == pc and wstyle == pstyle:
                matched.append({
                    "Style": wstyle, "Style 2": pstyle,
                    "WO Size": ws, "PO Size": ps,
                    "WO Colour Code": wc, "PO Colour Code": pc,
                    "WO Qty": wq, "PO Qty": pq,
                    "Qty Match": "👍 Yes", "Size Match": "👍 Yes",
                    "Colour Match": "👍 Yes", "Style Match": "👍 Yes",
                    "Diff": 0,
                    "Status": "🟩 Full Match",
                    "PO Item Code": po.get("Item_Code", "")
                })
                used.add(idx)
                full_match_found = True
                break
            else:
                score = 0
                if abs(pq - wq) <= tolerance:
                    score += 1
                if ws == ps:
                    score += 1
                if wc == pc:
                    score += 1
                if wstyle == pstyle:
                    score += 1
                if score > partial_match_score:
                    partial_match_score = score
                    partial_match_idx = idx
        if full_match_found:
            continue
        if partial_match_idx is not None:
            po = po_details[partial_match_idx]
            pq = po["Quantity"]
            ps = po.get("Size", "").strip().upper()
            pc = po.get("Colour_Code", "").strip().upper()
            pstyle = po.get("Style 2", "").strip()
            qty_match = "👍 Yes" if abs(pq - wq) <= tolerance else "❌ No"
            size_match = "👍 Yes" if ws == ps else "❌ No"
            colour_match = "👍 Yes" if wc == pc else "❌ No"
            style_match = "👍 Yes" if wstyle == pstyle else "❌ No"
            matched.append({
                "Style": wstyle, "Style 2": pstyle,
                "WO Size": ws, "PO Size": ps,
                "WO Colour Code": wc, "PO Colour Code": pc,
                "WO Qty": wq, "PO Qty": pq,
                "Qty Match": qty_match, "Size Match": size_match,
                "Colour Match": colour_match, "Style Match": style_match,
                "Diff": pq - wq,
                "Status": "🟨 Partial Match",
                "PO Item Code": po.get("Item_Code", "")
            })
            used.add(partial_match_idx)
        else:
            mismatched.append({
                "Style": wstyle, "Style 2": "",
                "WO Size": ws, "PO Size": "",
                "WO Colour Code": wc, "PO Colour Code": "",
                "WO Qty": wq, "PO Qty": None,
                "Qty Match": "❌ No", "Size Match": "❌ No",
                "Colour Match": "❌ No", "Style Match": "❌ No",
                "Diff": "", "Status": "❌ No PO Match", "PO Item Code": ""
            })
    for idx, po in enumerate(po_details):
        if idx not in used:
            ps = po.get("Size", "").strip().upper()
            pc = po.get("Colour_Code", "").strip().upper()
            pstyle = po.get("Style 2", "").strip()
            pq = po["Quantity"]
            mismatched.append({
                "Style": "", "Style 2": pstyle,
                "WO Size": "", "PO Size": ps,
                "WO Colour Code": "", "PO Colour Code": pc,
                "WO Qty": None, "PO Qty": pq,
                "Qty Match": "❌ No", "Size Match": "❌ No",
                "Colour Match": "❌ No", "Style Match": "❌ No",
                "Diff": "", "Status": "❌ Extra PO Item", "PO Item Code": po.get("Item_Code", "")
            })
    matched = sort_items_by_size(matched)
    mismatched = sort_items_by_size(mismatched)
    return matched, mismatched

def compare_codes(po_details, wo_items):
    po_codes = set()
    for po in po_details:
        code = po.get("Product_Code", "")
        if code:
            if isinstance(code, list):
                code_str = " / ".join(str(c) for c in code if c)
            else:
                code_str = str(code)
            cleaned_code = code_str.strip().upper()
            if cleaned_code:
                po_codes.add(cleaned_code)
    wo_codes = set()
    for w in wo_items:
        code = w.get("WO Product Code", "")
        if code:
            if isinstance(code, list):
                code_str = " / ".join(str(c) for c in code if c)
            else:
                code_str = str(code)
            cleaned_code = code_str.strip().upper()
            if cleaned_code:
                wo_codes.add(cleaned_code)
    comparison = []
    all_codes = po_codes.union(wo_codes)
    for code in all_codes:
        in_po = code in po_codes
        in_wo = code in wo_codes
        status = "✅ Match" if in_po and in_wo else "❌ Missing in WO" if in_po else "❌ Missing in PO"
        comparison.append({"PO Code": code if in_po else "", "WO Code": code if in_wo else "", "Status": status})
    return comparison

# -------------------- Debug Function --------------------
def debug_po_extraction(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    lines = [ln.strip() for ln in text.split("\n")]
    delivery_location_index = -1
    for i, line in enumerate(lines):
        if "Delivery Location:" in line:
            delivery_location_index = i
            break
    st.write(f"Found 'Delivery Location:' at line {delivery_location_index}")
    if delivery_location_index != -1:
        st.write("Next 10 lines after 'Delivery Location:':")
        for i in range(delivery_location_index + 1, min(delivery_location_index + 11, len(lines))):
            st.write(f"Line {i}: {lines[i]}")
    capture = False
    address_lines = []
    for ln in lines:
        if "Delivery Location:" in ln:
            capture = True
            continue
        if capture:
            if "Forwarder:" in ln:
                break
            if not ln:
                break
            address_lines.append(ln)
    full_address = " ".join(address_lines)
    st.write("### Full Address Text")
    st.write(full_address)
    plot_patterns = [
        "plot #", "plot no", "plot no.", "plot number", "plot",
        "building #", "building no", "building no.", "building number",
        "door #", "door no", "door no.", "door number"
    ]
    st.write("### Pattern Matching")
    found_patterns = []
    for pattern in plot_patterns:
        if pattern.lower() in full_address.lower():
            found_patterns.append(pattern)
            st.write(f"✅ Found pattern: {pattern}")
        else:
            st.write(f"❌ Pattern not found: {pattern}")
    india_count = full_address.lower().count("india")
    st.write(f"### 'India' Occurrences: {india_count}")
    po_fields = extract_po_fields(pdf_file)
    st.write("### Extracted Address")
    st.write(po_fields["delivery_location"])
    return None

# -------------------- Sidebar Configuration --------------------
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 1rem 0; color: black;">
        <h2>⚙️ Control Panel</h2>
        <p style="opacity: 0.8;">Configure your analysis settings</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📁 File Upload")
    wo_file = st.file_uploader(
        "📄 Work Order (WO) PDF", 
        type="pdf",
        disabled=not selected_user,
        help="Upload your Work Order PDF file"
    )
    
    po_file = st.file_uploader(
        "📋 Purchase Order (PO) PDF", 
        type="pdf",
        disabled=not selected_user,
        help="Upload your Purchase Order PDF file"
    )
    
    if wo_file:
        st.success("✅ WO File Loaded")
    if po_file:
        st.success("✅ PO File Loaded")
    
    if wo_file and po_file:
        st.markdown("### 🚀 Ready to Process")
        st.info("Both files are loaded. Analysis will begin automatically.")

# -------------------- Excel/PDF Merger Section --------------------
with st.expander("📓 Excel Table Data Extractor", expanded=False):
    st.markdown("""
    <div class="section-header">
        <h3 class="section-title">📊 Excel Table Data Extractor</h3>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        excel_file = st.file_uploader(
            "📤 Upload Excel File", 
            type=["xls", "xlsx"], 
            key="excel_merger",
            disabled=not selected_user,
            help="Upload Excel file containing table data starting at row 22"
        )
    
    with col2:
        pdf_file_merger = st.file_uploader(
            "📄 Upload PDF File (Optional)", 
            type=["pdf"], 
            key="pdf_merger",
            disabled=not selected_user,
            help="Optional: Upload PDF file to merge with extracted styles"
        )
    
    if excel_file:
        with st.spinner("🔄 Extracting table data..."):
            # Extract table data from all sheets
            all_sheets_data = read_excel_table(excel_file)
            
            # Process the extracted data
            if all_sheets_data:
                # Filter out sheets with no data
                sheets_with_data = [s for s in all_sheets_data if not s['data'].empty]
                
                if sheets_with_data:
                    processed_data = process_excel_table_data(sheets_with_data)
                    
                    # Store processed data in session state
                    st.session_state.processed_excel_data = processed_data
                    
                    # Extract style numbers for potential PDF merging
                    styles = []
                    for sheet_data in sheets_with_data:
                        if sheet_data['style_number']:
                            styles.append(sheet_data['style_number'])
                    
                    # If no style numbers found in STYLE column, try to extract from the data
                    if not styles:
                        for sheet_data in sheets_with_data:
                            df = sheet_data['data']
                            if not df.empty:
                                # Look for any 8-digit number in the dataframe
                                for col in df.columns:
                                    for val in df[col].dropna():
                                        val_str = str(val).strip()
                                        if re.match(r'^\d{8}$', val_str):
                                            styles.append(val_str)
                                            break
                                    if styles:
                                        break
                                if styles:
                                    break
                    
                    # Display the extracted table data
                    st.markdown("""
                    <div class="alert-success">
                        ✅ <strong>Success!</strong> Table data extracted successfully.
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Show extraction details
                    st.markdown("### 📋 Extraction Details")
                    st.markdown("""
                    <div class="alert-info">
                        ℹ️ <strong>Extraction Rules:</strong> 
                        <ul>
                            <li>Starts at row 22</li>
                            <li>Skips row 23</li>
                            <li>Skips any row with stopping text</li>
                            <li>Removes all unnamed columns</li>
                            <li>Stops at first blank row in STYLE column</li>
                            <li>Only reads columns up to QTY column</li>
                        </ul>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    for sheet_data in all_sheets_data:
                        with st.expander(f"Sheet: {sheet_data['sheet_name']}"):
                            st.markdown(f"**Style Number:** {sheet_data['style_number'] or 'Not found'}")
                            
                            # Display stopping information
                            if sheet_data['stop_row'] is not None:
                                if sheet_data['stop_row'] == 21:
                                    st.markdown(f"""
                                    <div class="alert-warning">
                                        ⚠️ <strong>Stopping text found in header row (row 22).</strong> No data extracted from this sheet.
                                    </div>
                                    """, unsafe_allow_html=True)
                                else:
                                    st.markdown(f"""
                                    <div class="alert-warning">
                                        ⚠️ <strong>Stopping text found at row {sheet_data['stop_row'] + 1}.</strong> This row and any subsequent rows with stopping text were skipped.
                                    </div>
                                    """, unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="alert-success">
                                    ✅ <strong>No stopping text found.</strong> All rows (except row 23) were processed.
                                </div>
                                """, unsafe_allow_html=True)
                            
                            # Display QTY column information
                            if sheet_data['qty_col_idx'] is not None:
                                st.markdown(f"**QTY column found at index:** {sheet_data['qty_col_idx']}")
                            else:
                                st.markdown("""
                                <div class="alert-warning">
                                    ⚠️ <strong>QTY column not found.</strong> All columns were read.
                                </div>
                                """, unsafe_allow_html=True)
                            
                            # Display the data if it exists
                            if not sheet_data['data'].empty:
                                # Check if data was truncated due to blank STYLE
                                if 'STYLE' in sheet_data['data'].columns:
                                    last_style = sheet_data['data']['STYLE'].iloc[-1]
                                    if pd.isna(last_style) or str(last_style).strip() == '':
                                        st.markdown("""
                                        <div class="alert-info">
                                            ℹ️ <strong>Data truncated</strong> at first blank row in STYLE column.
                                        </div>
                                        """, unsafe_allow_html=True)
                                
                                st.markdown("**Extracted Data:**")
                                st.dataframe(sheet_data['data'], use_container_width=True)
                                
                                # Show removed unnamed columns info
                                st.markdown("""
                                <div class="alert-info">
                                    ℹ️ <strong>Note:</strong> All unnamed columns have been removed.
                                </div>
                                """, unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                <div class="alert-info">
                                    ℹ️ <strong>No data extracted</strong> from this sheet.
                                </div>
                                """, unsafe_allow_html=True)
                    
                    # Show processed combined data
                    st.markdown("### 🔄 Processed Combined Data")
                    st.dataframe(processed_data, use_container_width=True, hide_index=True)
                    
                    # Download options
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Download processed data as Excel
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            processed_data.to_excel(writer, index=False, sheet_name='Processed Data')
                        output.seek(0)
                        
                        st.download_button(
                            label="⬇️ Download Processed Excel",
                            data=output,
                            file_name="Processed_Excel_Data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                    with col2:
                        # If styles were found and PDF is uploaded, offer merged PDF
                        if styles and pdf_file_merger:
                            styles_pdf = create_styles_pdf(styles)
                            final_pdf = merge_pdfs(pdf_file_merger, styles_pdf)
                            
                            st.download_button(
                                label="⬇️ Download Merged PDF",
                                data=final_pdf,
                                file_name="Merged-PO.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )
                        elif styles and not pdf_file_merger:
                            st.info("Upload a PDF file to merge with extracted styles")
                else:
                    st.markdown("""
                    <div class="alert-warning">
                        ⚠️ <strong>Warning:</strong> No valid table data found in any sheet.
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div class="alert-warning">
                    ⚠️ <strong>Warning:</strong> No valid table data found in the Excel file.
                </div>
                """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="alert-info">
            ℹ️ <strong>Info:</strong> Please upload an Excel file to extract table data.
        </div>
        """, unsafe_allow_html=True)

# -------------------- Main Analysis Section --------------------
if selected_user and wo_file and po_file:
    with st.expander("🔍 Debug PO Address Extraction"):
        st.write("### PO Address Extraction Debug Information")
        debug_po_extraction(po_file)
    
    st.markdown(show_progress_steps(2), unsafe_allow_html=True)
    
    with st.spinner("🔄 Processing files and analyzing data..."):
        wo = extract_wo_fields(wo_file)
        po = extract_po_fields(po_file)
        wo_items = extract_wo_items_table(wo_file, wo["product_codes"])
        wo_items = reorder_wo_by_size(wo_items)
        po_details_raw = extract_po_details(po_file)
        po_details = reorder_po_by_size(po_details_raw)
        addr_res = compare_addresses(wo, po)
        code_res = compare_codes(po_details, wo_items)
        matched, mismatched = enhanced_quantity_matching(wo_items, po_details)
        po_number = extract_po_number(po_file)
    
    st.markdown(show_progress_steps(4), unsafe_allow_html=True)
    
    st.markdown("""
    <div class="alert-success">
        🎉 <strong>Analysis Complete!</strong> Your files have been processed successfully.
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="section-header">
        <h2 class="section-title">📊 Analysis Overview</h2>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_wo_items = len(wo_items)
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total_wo_items}</div>
            <div class="metric-label">WO Items</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        total_po_items = len(po_details)
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total_po_items}</div>
            <div class="metric-label">PO Items</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        total_matched = len([m for m in matched if "Full Match" in m.get("Status", "")])
        st.markdown(f"""
        <div class="metric-card" style="border-top-color: #28a745;">
            <div class="metric-value" style="color: #28a745;">{total_matched}</div>
            <div class="metric-label">Perfect Matches</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        total_mismatched = len(mismatched)
        st.markdown(f"""
        <div class="metric-card" style="border-top-color: #dc3545;">
            <div class="metric-value" style="color: #dc3545;">{total_mismatched}</div>
            <div class="metric-label">Mismatches</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="section-header">
        <h3 class="section-title">🏠 Address Verification</h3>
    </div>
    """, unsafe_allow_html=True)
    
    addr_df = pd.DataFrame([addr_res])
    st.dataframe(addr_df, use_container_width=True, hide_index=True)
    
    if addr_res.get("Status") == "✅ Match":
        st.markdown("""
        <div class="alert-success">
            ✅ <strong>Address Match:</strong> Delivery addresses are verified and match successfully.
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="alert-warning">
            ⚠️ <strong>Address Review Required:</strong> Please verify the delivery addresses manually.
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="section-header">
        <h3 class="section-title">🔢 Product Code Analysis</h3>
    </div>
    """, unsafe_allow_html=True)
    
    po_all_codes = [po.get("Product_Code", "").strip().upper() for po in po_details if po.get("Product_Code")]
    wo_all_codes = [wo.get("WO Product Code", "").strip().upper() for wo in wo_items if wo.get("WO Product Code")]
    comparison_rows = []
    max_len = max(len(po_all_codes), len(wo_all_codes)) if po_all_codes or wo_all_codes else 0
    for i in range(max_len):
        po_code = po_all_codes[i] if i < len(po_all_codes) else ""
        wo_code = wo_all_codes[i] if i < len(wo_all_codes) else ""
        if po_code and wo_code and po_code == wo_code:
            status = "✅ Exact Match"
        elif po_code and wo_code and "/" in wo_code:
            wo_parts = [part.strip().upper() for part in wo_code.split("/")]
            status = "✅ Exact Match" if po_code in wo_parts else "❌ No Match"
        elif po_code and wo_code and "/" in po_code:
            po_parts = [part.strip().upper() for part in po_code.split("/")]
            status = "✅ Partial Match" if wo_code in po_parts else "❌ No Match"
        elif po_code and wo_code:
            status = "❌ No Match"
        else:
            status = "⚪ Empty"
        comparison_rows.append({
            "📋 PO Product Code": po_code,
            "📄 WO Product Code": wo_code,
            "🔍 Match Status": status
        })
    code_table_df = pd.DataFrame(comparison_rows)
    st.dataframe(code_table_df, use_container_width=True, hide_index=True)
    
    st.markdown("""
    <div class="section-header">
        <h3 class="section-title">📋 Matching other items</h3>
    </div>
    """, unsafe_allow_html=True)
    
    if matched:
        matched_df = pd.DataFrame(matched)
        st.dataframe(matched_df, use_container_width=True, hide_index=True)
        perfect_matches = len([m for m in matched if "Full Match" in m.get("Status", "")])
        if perfect_matches == len(matched):
            st.markdown("""
            <div class="alert-success">
                🎯 <strong>Perfect Score!</strong> All matched items have complete data alignment.
            </div>
            """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="alert-info">
            ℹ️ <strong>No Matches Found:</strong> No items were matched with the current algorithm.
        </div>
        """, unsafe_allow_html=True)
    
    # New section: Additional Data
    st.markdown("""
    <div class="section-header">
        <h3 class="section-title">📊 Additional Data</h3>
    </div>
    """, unsafe_allow_html=True)
    
    # Display WO Items table
    st.markdown("### 📄 Work Order (WO) Items")
    wo_df = pd.DataFrame(wo_items)
    st.dataframe(wo_df, use_container_width=True, hide_index=True)
    
    # Display Processed Combined Data from Excel extraction right after WO Items
    st.markdown("### 📊 Processed Combined Data from Excel")
    if hasattr(st.session_state, 'processed_excel_data') and st.session_state.processed_excel_data is not None:
        st.dataframe(st.session_state.processed_excel_data, use_container_width=True, hide_index=True)
    else:
        st.markdown("""
        <div class="alert-info">
            ℹ️ <strong>Info:</strong> No Excel data processed yet. Please upload and process an Excel file in the "Excel Table Data Extractor" section.
        </div>
        """, unsafe_allow_html=True)
    
    address_ok = addr_res.get("Status", "") == "✅ Match"
    codes_ok = not code_table_df.empty and all(code_table_df["🔍 Match Status"].isin(["✅ Exact Match", "✅ Partial Match"]))
    matched_df = pd.DataFrame(matched) if matched else pd.DataFrame()
    matched_ok = not matched_df.empty and all(matched_df["Status"] == "🟩 Full Match")
    mismatched_empty = len(mismatched) == 0
    
    if address_ok and codes_ok and matched_ok and mismatched_empty:
        match_status = "PERFECT MATCH!"
        st.markdown("""
        <audio autoplay>
            <source src="https://assets.mixkit.co/sfx/preview/mixkit-winning-chimes-2015.mp3" type="audio/mpeg">
        </audio>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class="alert-success" style="text-align: center; font-size: 1.2rem;">
            🎉 <strong>PERFECT MATCH!</strong> All verification checks passed successfully! 🎉
        </div>
        """, unsafe_allow_html=True)
        
        st.balloons()
        st.balloons()
        
    else:
        match_status = "NOT PERFECT"
        st.markdown("""
        <div class="alert-warning">
            ⚠️ <strong>Review Required:</strong> Some data points need manual verification. Check the details below.
        </div>
        """, unsafe_allow_html=True)
    
    wo_product_codes = []
    for item in wo_items:
        code = item.get("WO Product Code", "")
        if code:
            if isinstance(code, list):
                for c in code:
                    if c and c.strip():
                        wo_product_codes.append(c.strip().upper())
            elif code.strip():
                wo_product_codes.append(code.strip().upper())
    
    references = []
    extracted_styles = extract_style_numbers_from_po_first_page(po_file)
    if extracted_styles:
        references.extend(extracted_styles)
    for item in wo_items:
        style = item.get("Style", "")
        if style and style not in references:
            references.append(style)
    for item in po_details:
        style = item.get("Style 2", "")
        if style and style not in references:
            references.append(style)
    
    first_product_code = wo_product_codes[0] if wo_product_codes else ""
    first_reference = references[0] if references else ""
    
    with st.spinner("📊 Logging report to text file..."):
        success, message = log_to_text(selected_user, first_product_code, first_reference, match_status, po_number)
        if success:
            st.success(message)
        else:
            st.error(message)
    
    st.markdown("""
    <div class="section-header">
        <h3 class="section-title">❗ Mismatch Summary</h3>
    </div>
    """, unsafe_allow_html=True)
    
    if mismatched:
        st.markdown(f"""
        <div class="mismatch-summary">
            <h4>⚠️ Mismatch Detected</h4>
            <p>Found <strong>{len(mismatched)} mismatched items</strong> requiring attention.</p>
            <p>Below is an example of one mismatched item:</p>
        </div>
        """, unsafe_allow_html=True)
        first_mismatched = mismatched[0]
        mismatch_df = pd.DataFrame([first_mismatched])
        st.markdown("""
        <div class="mismatch-example">
            <h5>Example Mismatched Item:</h5>
        </div>
        """, unsafe_allow_html=True)
        st.dataframe(mismatch_df, use_container_width=True, hide_index=True)
        with st.expander("View All Mismatched Items"):
            all_mismatched_df = pd.DataFrame(mismatched)
            st.dataframe(all_mismatched_df, use_container_width=True, hide_index=True)
    else:
        st.markdown("""
        <div class="alert-success">
            ✅ <strong>No Mismatches!</strong> All items have been successfully matched.
        </div>
        """, unsafe_allow_html=True)
    
    with st.expander("📊 Detailed Data Tables", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### 📄 Work Order (WO) Items")
            wo_df = pd.DataFrame(wo_items)
            st.dataframe(wo_df, use_container_width=True, hide_index=True)
        with col2:
            st.markdown("### 📋 Purchase Order (PO) Items")
            po_df = pd.DataFrame(po_details)
            st.dataframe(po_df, use_container_width=True, hide_index=True)

# -------------------- Footer --------------------
st.markdown("""
<div class="footer">
    <p>
        🚀 <strong>Customer Care System v2.0</strong> | 
        Powered by <strong>Razz....</strong> | 
        Advanced PDF Analysis & Comparison Technology
    </p>
</div>
""", unsafe_allow_html=True)