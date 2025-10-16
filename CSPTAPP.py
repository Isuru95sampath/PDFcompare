import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import os
import re

# Import from other modules
try:
    from po_extraction import extract_po_fields, extract_po_details, extract_po_number, extract_style_numbers_from_po_first_page
    from wo_extraction import extract_wo_fields, extract_wo_items_table, reorder_wo_by_size
    from comparison_functions import compare_addresses, compare_codes, enhanced_quantity_matching, sort_items_by_size
    from excel_extraction import read_excel_table, process_excel_table_data, convert_excel_size_codes, combine_wo_and_excel_data
except ImportError as e:
    st.error(f"Import error: {e}")
    st.error("Please ensure all module files are in the same directory as the main application.")
    st.stop()
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

# -------------------- Style to PDF --------------------
def create_styles_pdf(styles: list) -> BytesIO:
    import fitz  # PyMuPDF
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
    import fitz  # PyMuPDF
    pdf_out = fitz.open()
    pdf_styles = fitz.open(stream=styles_pdf.read(), filetype="pdf")
    pdf_orig = fitz.open(stream=original_pdf.read(), filetype="pdf")
    
    pdf_out.insert_pdf(pdf_styles)
    pdf_out.insert_pdf(pdf_orig)
    
    output = BytesIO()
    pdf_out.save(output)
    output.seek(0)
    return output

# -------------------- Pattern --------------------
style_pattern = re.compile(r"^\d{8}$")

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
    
    # New section: Combined Data
    st.markdown("""
    <div class="section-header">
        <h3 class="section-title">📊 Combined WO and Excel Data</h3>
    </div>
    """, unsafe_allow_html=True)
    
    # Prepare WO Items table with empty columns removed and WO Product Code removed
    wo_df = pd.DataFrame(wo_items)
    for col in wo_df.columns:
        if wo_df[col].isnull().all() or (wo_df[col].astype(str).str.strip() == '').all():
            wo_df = wo_df.drop(columns=[col])
    
    # Remove WO Product Code column if it exists
    if 'WO Product Code' in wo_df.columns:
        wo_df = wo_df.drop(columns=['WO Product Code'])
    
    # Check if we have Excel data
    if hasattr(st.session_state, 'processed_excel_data') and st.session_state.processed_excel_data is not None:
        excel_df = st.session_state.processed_excel_data
        
        if not excel_df.empty:
            # Combine the data
            combined_df = combine_wo_and_excel_data(wo_df, excel_df)
            
            # Display the combined table
            st.dataframe(combined_df, use_container_width=True, hide_index=True)
            
            # Add download button for combined data
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False, sheet_name='Combined Data')
            output.seek(0)
            
            st.download_button(
                label="⬇️ Download Combined Data",
                data=output,
                file_name="Combined_WO_Excel_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.markdown("""
            <div class="alert-warning">
                ⚠️ <strong>Empty Data:</strong> The processed Excel data is empty.
            </div>
            """, unsafe_allow_html=True)
            st.dataframe(wo_df, use_container_width=True, hide_index=True)
    else:
        st.markdown("""
        <div class="alert-info">
            ℹ️ <strong>Info:</strong> No Excel data processed yet. Please upload and process an Excel file in the "Excel Table Data Extractor" section to see the combined table.
        </div>
        """, unsafe_allow_html=True)
        st.dataframe(wo_df, use_container_width=True, hide_index=True)
    
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
            wo_df_detailed = pd.DataFrame(wo_items)
            for col in wo_df_detailed.columns:
                if wo_df_detailed[col].isnull().all() or (wo_df_detailed[col].astype(str).str.strip() == '').all():
                    wo_df_detailed = wo_df_detailed.drop(columns=[col])
            # Remove WO Product Code column if it exists
            if 'WO Product Code' in wo_df_detailed.columns:
                wo_df_detailed = wo_df_detailed.drop(columns=['WO Product Code'])
            st.dataframe(wo_df_detailed, use_container_width=True, hide_index=True)
        with col2:
            st.markdown("### 📋 Purchase Order (PO) Items")
            po_df = pd.DataFrame(po_details)
            st.dataframe(po_df, use_container_width=True, hide_index=True)

# -------------------- Debug Function --------------------
def debug_po_extraction(pdf_file):
    import pdfplumber
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