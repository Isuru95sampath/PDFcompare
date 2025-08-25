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
    page_title="CS Data Entry Checking Tool – Price Tickets - Razz Solutions",
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
</style>
""", unsafe_allow_html=True)

# -------------------- Main Header --------------------
st.markdown("""
<div class="main-header">
    <h1 class="main-title">🚀 CS Data Entry Checking Tool – Price Tickets</h1>
    <p class="main-subtitle">Advanced PO vs WO Comparison Dashboard | Powered by Razz </p>
</div>
""", unsafe_allow_html=True)

# -------------------- Function to log to Excel --------------------
def log_to_excel(username, product_code, style_number, match_status):
    """Log analysis results to Excel file - one row per analysis"""
    try:
        # Define the path to the Excel file
        excel_path = r"C:\Users\Pcadmin\Documents\ITL\PriceTicket\CS-AI tool report.xlsx"
        
        # Create the directory if it doesn't exist
        os.makedirs(os.path.dirname(excel_path), exist_ok=True)
        
        # Get current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Prepare the row data - one row per analysis
        row_to_insert = [
            timestamp,
            username,
            product_code,
            style_number,
            match_status
        ]
        
        # Check if the file exists
        if os.path.exists(excel_path):
            # Load the existing workbook
            wb = openpyxl.load_workbook(excel_path)
            ws = wb.active
        else:
            # Create a new workbook
            wb = openpyxl.Workbook()
            ws = wb.active
            # Write header
            ws.append(["Timestamp", "Username", "Product Code", "Style Number", "Match Status"])
        
        # Append the row
        ws.append(row_to_insert)
        
        # Save the workbook
        wb.save(excel_path)
        
        return True, "Report successfully logged to Excel"
    except Exception as e:
        return False, f"Error logging to Excel: {e}"

# -------------------- Function to View Excel Data --------------------
def view_excel_data():
    """View data from the Excel file"""
    try:
        # Define the path to the Excel file
        excel_path = r"C:\Users\Pcadmin\Documents\ITL\PriceTicket\CS-AI tool report.xlsx"
        
        # Check if the file exists
        if os.path.exists(excel_path):
            # Read the Excel file
            df = pd.read_excel(excel_path)
            return df
        else:
            return pd.DataFrame(columns=["Timestamp", "Username", "Product Code", "Style Number", "Match Status"])
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return pd.DataFrame()

# -------------------- User Selection --------------------
with st.sidebar:
    st.markdown("### 👤 User Selection")
    selected_user = st.selectbox(
        "Select your name:",
        ["", "saman", "kasun", "gayan", "nipun", "udara"],
        help="You must select a user to access the application"
    )
    
    if not selected_user:
        st.warning("⚠️ Please select a user to continue")
    
    # Excel management section
    st.markdown("### 📊 Excel Management")
    if st.button("View All Data"):
        data_df = view_excel_data()
        if not data_df.empty:
            st.dataframe(data_df, use_container_width=True)
        else:
            st.info("No data available in the Excel file")

# -------------------- Progress Steps --------------------
def show_progress_steps(current_step=1):
    return ""

# -------------------- Pattern --------------------
style_pattern = re.compile(r"^\d{8}$")

# -------------------- Excel Reader --------------------
def extract_styles_from_excel(file) -> list:
    try:
        if file.name.endswith(".xls"):
            df = pd.read_excel(file, sheet_name=0, header=None, engine="xlrd")
            for i in range(23, len(df)):
                val = df.iloc[i, 1]
                val_str = str(val).strip() if val is not None else ""
                if not val_str:
                    break
                if style_pattern.match(val_str):
                    return [val_str]
                else:
                    break
        else:
            wb = openpyxl.load_workbook(file, data_only=True)
            ws = wb.active
            row = 24
            while True:
                val = ws[f"B{row}"].value
                val_str = str(val).strip() if val is not None else ""
                if not val_str:
                    break
                if style_pattern.match(val_str):
                    return [val_str]
                else:
                    break
                row += 1
        return []
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {e}")
        return []

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

# -------------------- Helper Functions --------------------
def clean_quantity(qty_str):
    """Convert strings like '1,148.00' or '465.00' into integers"""
    if not qty_str:
        return 0
    qty_str = str(qty_str).strip().replace(",", "")
    try:
        if "." in qty_str:
            return int(float(qty_str))
        return int(qty_str)
    except ValueError:
        return 0

def truncate_after_sri_lanka(addr: str) -> str:
    """
    Extract the address up to and including "Sri Lanka" or "India"
    """
    # First try to find "Sri Lanka"
    part, sep, _ = addr.partition("Sri Lanka")
    if sep:
        return (part + sep).strip()
    
    # If "Sri Lanka" not found, try to find "India"
    part, sep, _ = addr.partition("India")
    if sep:
        return (part + sep).strip()
    
    # If neither found, return the original address
    return addr.strip()

def clean_size(size_str):
    """Extract only the primary size from strings like 'S | P' or 'M / M'"""
    if not size_str:
        return ""
    size_str = str(size_str).strip().upper()
    
    # Split by pipe or slash and take the first part
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
            # If there's a line before "Deliver To:", check if it might be part of the address
            if i > 0:
                prev_line = lines[i-1].strip()
                # Check if the previous line looks like part of an address (not just a name)
                # If it contains numbers, street indicators, or is longer than a typical name
                if (any(char.isdigit() for char in prev_line) or 
                    any(indicator in prev_line.lower() for indicator in 
                        ["street", "st", "road", "rd", "avenue", "ave", "building", "block", "no", "#"]) or
                    len(prev_line.split()) > 3):  # More than 3 words suggests it's an address part
                    # Combine the previous line with the current address line
                    combined = prev_line + " " + re.sub(r"Deliver To:\s*", "", ln).strip()
                    delivery = truncate_after_sri_lanka(combined)
                else:
                    # Previous line is likely just a name, so skip it
                    delivery = truncate_after_sri_lanka(re.sub(r"Deliver To:\s*", "", ln).strip())
            else:
                # No previous line, just extract the address
                delivery = truncate_after_sri_lanka(re.sub(r"Deliver To:\s*", "", ln).strip())
            break
    
    # Remove "Customer Delivery Name:" if it exists in the address
    if "Customer Delivery Name:" in delivery:
        delivery = re.sub(r"Customer Delivery Name:\s*", "", delivery).strip()
    
    codes = []
    for line in lines:
        if "Product Code" in line:
            found = re.findall(r"Product Code[:\s]*([\w\s\-]+(?:\s*/\s*[\w\s\-]+)*)", line)
            for match in found:
                for code in match.split("/"):
                    clean = code.strip().upper()
                    
                    # New logic: Replace "-" with space
                    clean = clean.replace("-", " ")
                    
                    if clean:
                        codes.append(clean)
            break
    
    po_numbers = list(set(re.findall(r'\b\d{7,8}\b', text)))
    
    return {
        "customer_name": "",  # Empty as requested
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
    raw_addr = " ".join(address_lines)
    matches = re.findall(r".*Sri Lanka.*", text, re.IGNORECASE)
    unique = [raw_addr] + [m for m in matches if m != raw_addr]
    seen = []
    for a in unique:
        if a and a not in seen:
            seen.append(a)
    sri = [a for a in seen if "sri lanka" in a.lower()]
    chosen = max(sri, key=len) if sri else seen[0] if seen else raw_addr
    final_addr = truncate_after_sri_lanka(chosen)
    po_codes = re.findall(r"(LB\s*\d+)", text)
    sup_ref_codes = re.findall(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z]+[-\s]?\d+)", text, re.IGNORECASE)
    tag_codes = re.findall(r"TAG\.PRC\.TKT_(.*?)_REG", text)
    all_product_codes = list(set([c.strip().upper() for c in sup_ref_codes + tag_codes]))
    return {
        "delivery_location": final_addr,
        "product_codes": po_codes + all_product_codes,
        "all_found_addresses": seen
    }

def extract_style_numbers(po_pdf_path):
    style_numbers = set()
    with pdfplumber.open(po_pdf_path) as pdf:
        # Check PO tables
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        if row:
                            for cell in row:
                                if cell:
                                    cell_clean = str(cell).strip().replace(',', '')
                                    # If it's a decimal like "1148.00", take integer part
                                    if "." in cell_clean:
                                        cell_clean = cell_clean.split(".")[0]
                                    if cell_clean.isdigit():
                                        val = int(cell_clean)
                                        if 10 < val < 100000:  # filter out prices / too small / too big
                                            if val > qty:      # pick the largest number in the row
                                                qty = val
                                
        # Check free-text (regex search)
        full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        found_free_text_styles = re.findall(r'\b[A-Z]{2,}\s*\d{3,}\b', full_text)
        style_numbers.update([s.strip() for s in found_free_text_styles])
        # Check "Extracted Style Numbers:" section on first page
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
    """Sort matched or mismatched items by size from smallest to largest"""
    size_order = {"XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5}
    
    def get_size_key(item):
        # Use WO Size if available, otherwise use PO Size
        size = item.get("WO Size", item.get("PO Size", "")).strip().upper()
        return size_order.get(size, 99)  # Default to 99 if size not found
    
    return sorted(items, key=get_size_key)

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
        # NEW FORMAT HANDLING
        tag_match = re.search(r"TAG\.PRC\.TKT_(.*?)_REG", text)
        product_code_used = tag_match.group(1).strip().upper() if tag_match else ""
        
        product_code_used = product_code_used.replace("-", " ")
        
        i = 0
        while i < len(lines):
            line = lines[i]
            # Updated regex to handle quantities with commas
            item_match = re.match(r'^(\d+)\s+TAG\.PRC\.TKT_.*?([\d,]+\.\d+)\s+PCS', line)
            if item_match:
                item_no = item_match.group(1)
                quantity_str = item_match.group(2)
                quantity = clean_quantity(quantity_str)  # Using updated function
                
                colour = size = ""
                for j in range(i + 1, min(i + 5, len(lines))):
                    next_line = lines[j]
                    if "Color/Size/Destination :" in next_line:
                        cs_part = next_line.split(":", 1)[1].strip()
                        cs_parts = [part.strip() for part in cs_part.split(" / ") if part.strip()]
                        
                        if len(cs_parts) >= 2:
                            colour_part = cs_parts[0].strip()
                            colour = colour_part.split()[0] if colour_part else ""
                            size = cs_parts[1].strip().upper()
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
        # ORIGINAL FORMAT HANDLING (keep existing code, just ensure clean_quantity() is used)
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
            # Updated regex to handle quantities with commas
            item_match = re.match(r'^(\d+)\s+([A-Z0-9]+)\s+(\d+)\s+([\d,]+\.\d+)\s+PCS', line)
            if item_match:
                item_no, item_code, _, qty_str = item_match.groups()
                quantity = clean_quantity(qty_str)  # Using updated function
                colour = size = ""
                for j in range(i + 1, min(i + 10, len(lines))):
                    ln = lines[j]
                    if not colour and "Colour/Size/Destination:" in ln:
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
                item_key = (size.upper() if size else "", colour.upper() if colour else "", repeated_style)
                
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
    with improved table detection and data extraction for all formats
    """
    items = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # First, try to extract tables with standard settings
            tables = page.extract_tables()
            
            # If standard extraction doesn't work well, try with explicit settings
            if not tables or len(tables) == 0:
                tables = page.extract_tables({
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "explicit_vertical_lines": page.curves + page.edges,
                    "explicit_horizontal_lines": page.curves + page.edges,
                })
            
            # If still no tables, try another approach
            if not tables or len(tables) == 0:
                # Try to extract table by looking for text patterns
                text = page.extract_text()
                lines = text.split('\n')
                
                # Look for the header row
                header_idx = -1
                for i, line in enumerate(lines):
                    if "Style" in line and "Colour Code" in line and "Size" in line:
                        header_idx = i
                        break
                
                if header_idx >= 0:
                    # Process data rows
                    for i in range(header_idx + 1, len(lines)):
                        line = lines[i].strip()
                        if not line:
                            continue
                            
                        # Skip if it's a footer or separator
                        if "Number of Size Changes" in line or "End of Works Order" in line:
                            break
                            
                        # Try to extract data using regex - more flexible pattern
                        # This pattern is designed to capture all variations of sizes including XS, XL, XXL
                        pattern = r'(\d{8})\s+([A-Z0-9]+)\s+\b(XS|S|M|L|XL|XXL|XXXL|XXG|XG|P|G)\b.*?(\d{1,4}(?:,\d{3})*)'
                        match = re.search(pattern, line)
                        
                        if match:
                            style = match.group(1)
                            color_code = match.group(2)
                            size1 = match.group(3)
                            quantity_str = match.group(4).replace(',', '')
                            
                            try:
                                quantity = int(quantity_str)
                                items.append({
                                    "Style": style,
                                    "WO Colour Code": color_code.upper(),
                                    "Size 1": size1,
                                    "Quantity": quantity,
                                    "WO Product Code": " / ".join(product_codes) if product_codes else ""
                                })
                            except ValueError:
                                continue
            
            # Process tables if found
            for table in tables:
                if not table or len(table) < 2:
                    continue
                
                # Look for the header row
                header_row_idx = -1
                column_positions = {}
                
                for i, row in enumerate(table):
                    if not row:
                        continue
                    
                    row_text = " ".join([str(cell).strip() for cell in row if cell])
                    
                    # Check if this row contains expected headers
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
                
                # If no header found, try to identify by content patterns
                if header_row_idx == -1:
                    for i, row in enumerate(table):
                        if not row or len(row) < 8:
                            continue
                            
                        # Check if first column looks like a style number (8 digits)
                        first_cell = str(row[0]).strip()
                        if re.match(r'^\d{8}$', first_cell):
                            # Check if one of the columns contains size information
                            has_size = False
                            for cell in row:
                                cell_str = str(cell).strip().upper()
                                if any(size in cell_str for size in ["XS", "S", "M", "L", "XL", "XXL"]):
                                    has_size = True
                                    break
                            
                            if has_size:
                                header_row_idx = i
                                # Assume standard column positions
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
                                    "quantity": len(row) - 1  # Last column is quantity
                                }
                                break
                
                if header_row_idx == -1:
                    continue
                
                # Process data rows
                for row in table[header_row_idx + 1:]:
                    if not row or len(row) < max(column_positions.values()) + 1:
                        continue
                    
                    try:
                        # Extract data using column positions
                        style = str(row[column_positions.get("style", 0)] or "").strip()
                        color_code = str(row[column_positions.get("color_code", 1)] or "").strip().upper()
                        
                        # Extract and clean size1 - with enhanced detection for all sizes
                        size1_raw = str(row[column_positions.get("size1", 2)] or "").strip()
                        size1 = clean_size(size1_raw)
                        
                        # If size1 is still empty, try to find size in other columns
                        if not size1:
                            for cell in row:
                                cell_str = str(cell).strip().upper()
                                # Enhanced regex to capture all size variations
                                size_match = re.search(r'\b(XS|S|M|L|XL|XXL|XXXL|XXG|XG|P|G)\b', cell_str)
                                if size_match:
                                    size1 = size_match.group(1)
                                    break
                        
                        # Extract quantity - try multiple methods
                        quantity_str = ""
                        if "quantity" in column_positions:
                            quantity_str = str(row[column_positions["quantity"]] or "").strip()
                        
                        # If quantity not found or invalid, try the last column
                        if not quantity_str or not re.search(r'\d', quantity_str):
                            quantity_str = str(row[-1] or "").strip()
                        
                        # Clean and validate the quantity
                        quantity = clean_quantity(quantity_str)
                        
                        # Only add if we have valid data
                        if style and quantity > 0:
                            item_data = {
                                "Style": style,
                                "WO Colour Code": color_code,
                                "Size 1": size1,
                                "Quantity": quantity,
                                "WO Product Code": " / ".join(product_codes) if product_codes else ""
                            }
                            
                            # Add optional fields if available
                            if "size2" in column_positions:
                                size2_raw = str(row[column_positions["size2"]] or "").strip()
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
                        # Skip rows with parsing errors
                        continue
    
    # Enhanced fallback: If table extraction fails, try more comprehensive text pattern matching
    if not items:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        
        # Try multiple patterns to capture different table formats
        patterns = [
            # Pattern 1: Standard format with pipes - enhanced to capture all sizes
            r'(\d{8})\s+([A-Z0-9]{2,4})\s+\b(XS|S|M|L|XL|XXL|XXXL|XXG|XG|P|G)\b.*?(\d{1,4}(?:,\d{3})*)',
            # Pattern 2: Format without pipes but with clear column separation
            r'(\d{8})\s+([A-Z0-9]{2,4})\s+\b(XS|S|M|L|XL|XXL|XXXL|XXG|XG|P|G)\b.*?(\d{1,4}(?:,\d{3})*)',
            # Pattern 3: More flexible pattern for variations
            r'(\d{8})\s+([A-Z0-9]{2,4})\s+\b(XS|S|M|L|XL|XXL|XXXL|XXG|XG|P|G)\b.*?(\d{1,4}(?:,\d{3})*)',
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, full_text)
            if matches:
                for match in matches:
                    style_num, color_code, size, quantity = match
                    quantity = quantity.replace(',', '')
                    try:
                        items.append({
                            "Style": style_num,
                            "WO Colour Code": color_code.upper(),
                            "Size 1": size,
                            "Quantity": int(quantity),
                            "WO Product Code": " / ".join(product_codes) if product_codes else ""
                        })
                    except ValueError:
                        continue
                break  # Stop after first successful pattern match
    
    # Remove duplicates while preserving order
    seen = set()
    unique_items = []
    for item in items:
        key = (item["Style"], item["WO Colour Code"], item["Size 1"], item["Quantity"])
        if key not in seen:
            seen.add(key)
            unique_items.append(item)
    
    return unique_items

SIZES = {"XS", "S", "M", "L", "XL", "XXL", "XP", "P", "G", "XG"}

def _extract_qty_from_text(text, min_value=10, max_value=100000):
    """
    Find the most plausible quantity in a text cell.
    Handles: 1,148.00 -> 1148, 465.00 -> 465, 306.00 -> 306, 1675 -> 1675.
    Ignores small numbers by threshold and picks the largest candidate in the cell.
    """
    if not text:
        return None
    # Numbers like 1,148.00 or 1675 or 10.95
    candidates = re.findall(r'\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?', str(text))
    best = None
    for token in candidates:
        num_str = token.replace(",", "")
        try:
            # Take integer part (so 1148.00 -> 1148, 10.95 -> 10)
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
    """Remove duplicates while preserving order."""
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
    """
    Keep the same API as your 'enhanced' version. Uses the same logic above.
    """
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
    # Remaining unmatched PO items
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
    
    # Sort matched and mismatched items by size
    matched = sort_items_by_size(matched)
    mismatched = sort_items_by_size(mismatched)
    
    return matched, mismatched

def compare_addresses(wo, po):
    ns = fuzz.token_sort_ratio(wo["customer_name"], po["delivery_location"])
    as_ = fuzz.token_sort_ratio(wo["delivery_address"], po["delivery_location"])
    comb = max(ns, as_)
    return {"WO Name": wo["customer_name"], "WO Addr": wo["delivery_address"], "PO Addr": po["delivery_location"],
            "Name %": ns, "Addr %": as_, "Overall %": comb, "Status": "✅ Match" if comb > 85 else "⚠️ Review"}

def compare_codes(po_details, wo_items):
    # Handle PO codes - ensure they are strings before processing
    po_codes = set()
    for po in po_details:
        code = po.get("Product_Code", "")
        if code:
            # Convert to string if it's not already
            if isinstance(code, list):
                # Join list elements with a separator
                code_str = " / ".join(str(c) for c in code if c)
            else:
                code_str = str(code)
            
            # Clean and add to set
            cleaned_code = code_str.strip().upper()
            if cleaned_code:
                po_codes.add(cleaned_code)
    
    # Handle WO codes - ensure they are strings before processing
    wo_codes = set()
    for w in wo_items:
        code = w.get("WO Product Code", "")
        if code:
            # Convert to string if it's not already
            if isinstance(code, list):
                # Join list elements with a separator
                code_str = " / ".join(str(c) for c in code if c)
            else:
                code_str = str(code)
            
            # Clean and add to set
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

# -------------------- Sidebar Configuration --------------------
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 1rem 0; color: black;">
        <h2>⚙️ Control Panel</h2>
        <p style="opacity: 0.8;">Configure your analysis settings</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 🔧 Analysis Method")
    method = st.selectbox(
        "Select Matching Algorithm:",
        ["Enhanced Matching (with PO Color/Size)", "Smart Matching (Exact)", "Smart Matching with Tolerance"],
        help="Choose the comparison algorithm that best fits your data"
    )
    
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
    
    # File status indicators
    if wo_file:
        st.success("✅ WO File Loaded")
    if po_file:
        st.success("✅ PO File Loaded")
    
    if wo_file and po_file:
        st.markdown("### 🚀 Ready to Process")
        st.info("Both files are loaded. Analysis will begin automatically.")

# -------------------- Excel/PDF Merger Section --------------------
with st.expander("📓 PDF Style Number Merger", expanded=False):
    st.markdown("""
    <div class="section-header">
        <h3 class="section-title">📊 Excel to PDF Style Extractor</h3>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        excel_file = st.file_uploader(
            "📤 Upload Excel File", 
            type=["xls", "xlsx"], 
            key="excel",
            disabled=not selected_user,
            help="Upload Excel file containing style numbers"
        )
    
    with col2:
        pdf_file_merger = st.file_uploader(
            "📄 Upload PDF File", 
            type=["pdf"], 
            key="pdf_merger",
            disabled=not selected_user,
            help="Upload PDF file to merge with extracted styles"
        )
    if excel_file and pdf_file_merger:
        with st.spinner("🔄 Extracting styles and merging..."):
            styles = extract_styles_from_excel(excel_file)
        if styles:
            st.markdown("""
            <div class="alert-success">
                ✅ <strong>Success!</strong> Style numbers extracted successfully.
            </div>
            """, unsafe_allow_html=True)
            
            col1, col2 = st.columns([2, 1])
            with col1:
                st.dataframe(
                    pd.DataFrame(styles, columns=["📋 Extracted Style Numbers"]), 
                    use_container_width=True
                )
            
            with col2:
                styles_pdf = create_styles_pdf(styles)
                final_pdf = merge_pdfs(pdf_file_merger, styles_pdf)
                
                st.download_button(
                    label="⬇️ Download Merged PDF",
                    data=final_pdf,
                    file_name="Merged-PO.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
        else:
            st.markdown("""
            <div class="alert-warning">
                ⚠️ <strong>Warning:</strong> No valid style numbers found in the Excel file.
            </div>
            """, unsafe_allow_html=True)
    elif excel_file or pdf_file_merger:
        st.markdown("""
        <div class="alert-info">
            ℹ️ <strong>Info:</strong> Please upload both Excel and PDF files to proceed with merging.
        </div>
        """, unsafe_allow_html=True)

# -------------------- Main Analysis Section --------------------
if selected_user and wo_file and po_file:
    # Show progress
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
        if "Enhanced" in method:
            matched, mismatched = enhanced_quantity_matching(wo_items, po_details)
        else:
            matched, mismatched = [], []
    
    # Show completion
    st.markdown(show_progress_steps(4), unsafe_allow_html=True)
    
    # Success message
    st.markdown("""
    <div class="alert-success">
        🎉 <strong>Analysis Complete!</strong> Your files have been processed successfully.
    </div>
    """, unsafe_allow_html=True)
    
    # -------------------- Key Metrics Dashboard --------------------
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
    
    # -------------------- Address Comparison --------------------
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
    
    # -------------------- Product Code Comparison --------------------
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
    
    # -------------------- Item Matching Results --------------------
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
            ℹ️ <strong>No Matches Found:</strong> No items were matched with the current algorithm. Try adjusting the matching method.
        </div>
        """, unsafe_allow_html=True)
    
    # -------------------- Final Status Check --------------------
    address_ok = addr_res.get("Status", "") == "✅ Match"
    codes_ok = not code_table_df.empty and all(code_table_df["🔍 Match Status"].isin(["✅ Exact Match", "✅ Partial Match"]))
    matched_df = pd.DataFrame(matched) if matched else pd.DataFrame()
    matched_ok = not matched_df.empty and all(matched_df["Status"] == "🟩 Full Match")
    mismatched_empty = len(mismatched) == 0
    
    # Determine match status
    if address_ok and codes_ok and matched_ok and mismatched_empty:
        match_status = "PERFECT MATCH!"
        # Play success sound
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
    
    # -------------------- Collect Data for Excel Logging --------------------
    # Collect product codes
    wo_product_codes = []
    for item in wo_items:
        code = item.get("WO Product Code", "")
        if code:
            # Handle case where code might be a list
            if isinstance(code, list):
                for c in code:
                    if c and c.strip():
                        wo_product_codes.append(c.strip().upper())
            elif code.strip():
                wo_product_codes.append(code.strip().upper())
    
    # Collect style numbers
    style_numbers = []
    # Add style numbers from PO
    extracted_styles = extract_style_numbers_from_po_first_page(po_file)
    if extracted_styles:
        style_numbers.extend(extracted_styles)
    # Add style numbers from WO
    for item in wo_items:
        style = item.get("Style", "")
        if style and style not in style_numbers:
            style_numbers.append(style)
    # Add style numbers from PO details
    for item in po_details:
        style = item.get("Style 2", "")
        if style and style not in style_numbers:
            style_numbers.append(style)
    
    # Get first product code and first style number
    first_product_code = wo_product_codes[0] if wo_product_codes else ""
    first_style_number = style_numbers[0] if style_numbers else ""
    
    # Auto-log to Excel when files are uploaded
    with st.spinner("📊 Logging report to Excel..."):
        success, message = log_to_excel(selected_user, first_product_code, first_style_number, match_status)
        if success:
            st.success(message)
        else:
            st.error(message)
    
    # -------------------- Mismatched Items --------------------
    st.markdown("""
    <div class="section-header">
        <h3 class="section-title">❗ Items Requiring Attention</h3>
    </div>
    """, unsafe_allow_html=True)
    if mismatched:
        mismatched_df = pd.DataFrame(mismatched)
        st.dataframe(mismatched_df, use_container_width=True, hide_index=True)
        
        st.markdown(f"""
        <div class="alert-warning">
            ⚠️ <strong>{len(mismatched)} items</strong> require manual review or correction.
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="alert-success">
            ✅ <strong>No Mismatches!</strong> All items have been successfully matched.
        </div>
        """, unsafe_allow_html=True)
    
    # -------------------- Detailed Data Tables --------------------
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