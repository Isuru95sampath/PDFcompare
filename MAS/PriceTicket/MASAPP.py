import io
import re
import os
import pandas as pd
import streamlit as st
import pdfplumber

# Default Excel file path
DEFAULT_EXCEL_PATH = r"C:\Users\Pcadmin\Desktop\CP-SO-Tracker\CPEXCEL.xlsx"

# ---------------------- Helpers for Price Tickets ----------------------
def read_pdf_text(file_bytes: bytes) -> list[str]:
    texts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            text = re.sub(r"\u00A0", " ", text)
            texts.append(text)
    return texts

def full_text(pages: list[str]) -> str:
    return "\n".join(pages)

# ------------------ Field Extractors for Price Tickets ------------------
PO_NUM_PATTERNS = [
    r"\bPO\s*Number\s*[-:]*\s*(\d+)\b",
    r"\bPO\s*Number\s*\n\s*(\d+)\b",
]

def extract_po_number(text: str) -> str | None:
    for pat in PO_NUM_PATTERNS:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return None

def extract_product_codes(text: str) -> list[dict]:
    # Split text into lines for line-by-line processing
    lines = text.split('\n')
    result = []
    
    # Pattern to match item code and product code line for TKT
    tkt_pattern = r'^(\d+)\s+(TKT\s+.*)$'
    
    # Pattern to match terms and conditions section (numbered items)
    terms_pattern = r'^\d+\.\s+The\s+'
    
    # Pattern to match SO number
    so_pattern = r'^(\d+\s*/\s*\d+)$'
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Skip terms and conditions section
        if re.match(terms_pattern, line, re.IGNORECASE):
            # Skip all numbered items until we find a non-numbered line
            while i < len(lines) and re.match(r'^\d+\.', lines[i].strip()):
                i += 1
            continue
        
        # Try to match TKT pattern first
        tkt_match = re.match(tkt_pattern, line, re.IGNORECASE)
        
        if tkt_match:
            item_code = tkt_match.group(1).strip()
            full_product_code = tkt_match.group(2).strip()
            
            # Extract SO Number from the line immediately after this line (i+1)
            primary_so_number = None
            if i + 1 < len(lines):
                so_line = lines[i + 1].strip()
                primary_so_number = so_line
            
            # Remove "TKT" from the beginning of the product code
            without_tkt = re.sub(r'^TKT\s*', '', full_product_code, flags=re.IGNORECASE).strip()
            
            # Check if this is a TKT LB product code
            if "LB" in without_tkt:
                # Extract LB and 4 digits (skip any "-" between)
                lb_match = re.search(r'LB\s*-?(\d{4})', without_tkt, re.IGNORECASE)
                if lb_match:
                    product_code = "LB" + lb_match.group(1).strip()
                    
                    # Get the text after the product code
                    after_product = without_tkt[lb_match.end():].strip()
                    
                    # Find 8-digit style number
                    style_match = re.search(r'(\d{8})', after_product)
                    if style_match:
                        style = style_match.group(1)
                        
                        # Get the text after the style number
                        after_style = after_product[style_match.end():].strip()
                        
                        # Find 4-character code (letters and numbers) before slash
                        color_match = re.search(r'([A-Z0-9]{4})\s*/', after_style)
                        if color_match:
                            color_code = color_match.group(1)
                        else:
                            # Try to find 4-character code without slash
                            color_match = re.search(r'([A-Z0-9]{4})', after_style)
                            if color_match:
                                color_code = color_match.group(1)
                            else:
                                color_code = None
                    else:
                        style = None
                        color_code = None
                else:
                    product_code = without_tkt
                    style = None
                    color_code = None
            else:
                # Regular TKT product code processing
                # Extract base product code (up to and including first 'F')
                f_match = re.search(r"(.*?F)", without_tkt, re.IGNORECASE)
                base_product_code = f_match.group(1) if f_match else without_tkt
                
                # Get the text after the base product code
                after_base = without_tkt[len(base_product_code):].strip()
                
                # Find 8-digit style number
                style_match = re.search(r'(\d{8})', after_base)
                if style_match:
                    style = style_match.group(1)
                    
                    # Get the text after the style number
                    after_style = after_base[style_match.end():].strip()
                    
                    # Find 4-character code (letters and numbers) before slash
                    color_match = re.search(r'([A-Z0-9]{4})\s*/', after_style)
                    if color_match:
                        color_code = color_match.group(1)
                    else:
                        # Try to find 4-character code without slash
                        color_match = re.search(r'([A-Z0-9]{4})', after_style)
                        if color_match:
                            color_code = color_match.group(1)
                        else:
                            color_code = None
                else:
                    style = None
                    color_code = None
                
                product_code = base_product_code
            
            # Now extract table data belonging to this SO number
            table_data = []
            current_so_number = primary_so_number
            j = i + 2  # Start from the line after SO number
            
            # Continue until we hit another item code or end of document
            while j < len(lines):
                next_line = lines[j].strip()
                
                # Check if we've reached another item code
                if re.match(r'^\d+\s+(TKT|LB)', next_line, re.IGNORECASE):
                    break
                
                # Check if we've reached terms and conditions
                if re.match(terms_pattern, next_line, re.IGNORECASE):
                    break
                
                # Check if this line is an SO number
                so_match = re.match(so_pattern, next_line)
                if so_match:
                    # Update the current SO number
                    current_so_number = so_match.group(1)
                    j += 1
                    continue
                
                # Skip empty lines
                if not next_line:
                    j += 1
                    continue
                
                # Check if this line is a table row (starts with a number)
                if re.match(r'^\d', next_line):
                    tokens = next_line.split()
                    
                    # Extract table data if we have at least 3 tokens (line number, size, size qty)
                    if len(tokens) >= 3:
                        line_number = tokens[0]
                        size = tokens[1]
                        size_qty = tokens[2]
                        
                        table_data.append({
                            "Line Number": line_number,
                            "Size": size,
                            "Size Qty": size_qty,
                            "SO Number": current_so_number
                        })
                
                j += 1
            
            # Add all table data rows to the result
            for data in table_data:
                result.append({
                    "Item Code": item_code,
                    "Product Code": product_code,
                    "Style": style,
                    "Color Code": color_code,
                    "SO Number": data["SO Number"],
                    "Line Number": data["Line Number"],
                    "Size": data["Size"],
                    "Size Qty": data["Size Qty"]
                })
            
            # Move the index to the end of the current table data
            i = j
        else:
            # Check for LB product code pattern (without TKT)
            lb_match = re.match(r'^(\d+)\s+(LB\d{4})\s+(\d{8})\s+([A-Z0-9]{4})', line, re.IGNORECASE)
            if lb_match:
                item_code = lb_match.group(1).strip()
                product_code = lb_match.group(2).strip()  # LB followed by 4 digits
                style = lb_match.group(3).strip()          # 8 digits
                color_code = lb_match.group(4).strip()     # 4 alphanumeric characters
                
                # Extract SO Number from the line immediately after this line (i+1)
                primary_so_number = None
                if i + 1 < len(lines):
                    so_line = lines[i + 1].strip()
                    primary_so_number = so_line
                
                # Now extract table data belonging to this SO number
                table_data = []
                current_so_number = primary_so_number
                j = i + 2  # Start from the line after SO number
                
                # Continue until we hit another item code or end of document
                while j < len(lines):
                    next_line = lines[j].strip()
                    
                    # Check if we've reached another item code
                    if re.match(r'^\d+\s+(TKT|LB)', next_line, re.IGNORECASE):
                        break
                    
                    # Check if we've reached terms and conditions
                    if re.match(terms_pattern, next_line, re.IGNORECASE):
                        break
                    
                    # Check if this line is an SO number
                    so_match = re.match(so_pattern, next_line)
                    if so_match:
                        # Update the current SO number
                        current_so_number = so_match.group(1)
                        j += 1
                        continue
                    
                    # Skip empty lines
                    if not next_line:
                        j += 1
                        continue
                    
                    # Check if this line is a table row (starts with a number)
                    if re.match(r'^\d', next_line):
                        tokens = next_line.split()
                        
                        # Extract table data if we have at least 3 tokens (line number, size, size qty)
                        if len(tokens) >= 3:
                            line_number = tokens[0]
                            size = tokens[1]
                            size_qty = tokens[2]
                            
                            table_data.append({
                                "Line Number": line_number,
                                "Size": size,
                                "Size Qty": size_qty,
                                "SO Number": current_so_number
                            })
                    
                    j += 1
                
                # Add all table data rows to the result
                for data in table_data:
                    result.append({
                        "Item Code": item_code,
                        "Product Code": product_code,
                        "Style": style,
                        "Color Code": color_code,
                        "SO Number": data["SO Number"],
                        "Line Number": data["Line Number"],
                        "Size": data["Size"],
                        "Size Qty": data["Size Qty"]
                    })
                
                # Move the index to the end of the current table data
                i = j
            else:
                i += 1
    
    return result

# ---------------------- WO PDF Extractor Functions ----------------------
def extract_data_from_pdf(uploaded_file):
    """
    Extract data from MAS WO PDF with specific format handling for both formats
    """
    extracted_data = {
        "PO Number": [],
        "Item Code": [],
        "Product Code": [],
        "Style": [],
        "Color Code": [],
        "SO Number": [],
        "Line Number": [],
        "Size": [],
        "SKU Desc": [],   # ✅ Changed header name
        "Panty Length 2": [],
        "Retail (US)": [],
        "Retail (CA)": [],
        "Multi Price": [],
        "Product Desc": [],  # renamed for clarity (old SKU Desc was product description)
        "Article": [],
        "Quantity": []
    }

    try:
        with pdfplumber.open(uploaded_file) as pdf:
            text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"

        # Extract header info
        po_match = re.search(r'VS PO Number:\s*([^\n\r]+)', text, re.IGNORECASE)
        item_match = re.search(r'Item Code:\s*([^\n\r]+)', text, re.IGNORECASE)
        product_match = re.search(r'Product Code:\s*([^\n\r]+)', text, re.IGNORECASE)
        so_match = re.search(r'SO Number:\s*([^\n\r]+)', text, re.IGNORECASE)
        line_match = re.search(r'Line Item:\s*([^\n\r]+)', text, re.IGNORECASE)
        product_desc_match = re.search(r'Product Description:\s*([^\n\r]+)', text, re.IGNORECASE)

        po_number = po_match.group(1).strip() if po_match else ""
        item_code = item_match.group(1).strip() if item_match else ""
        product_code = product_match.group(1).strip() if product_match else ""
        so_number = so_match.group(1).strip() if so_match else ""
        line_number = line_match.group(1).strip() if line_match else ""
        product_desc = product_desc_match.group(1).strip() if product_desc_match else ""

        # Detect tables
        lines = text.split('\n')
        table_started = False
        table_rows = []
        table_format = None
        
        for i, line in enumerate(lines):
            line = line.strip()

            # Detect table header
            if "Style Colour Code Size Panty Length" in line:
                if "Retail (US)" in line and "Retail (CA)" in line:
                    table_format = 'extended'
                else:
                    table_format = 'basic'
                table_started = True
                continue

            if table_started:
                if ("Number of Size Changes" in line or 
                    "End of Works Order" in line or 
                    line.startswith("International Trimmings")):
                    break

                if not line:
                    continue

                if table_format == 'extended':
                    parts = re.split(r'\s+', line)
                    if len(parts) >= 8:
                        style = parts[0]
                        color_code = parts[1]

                        size_parts, price_start_idx = [], None
                        for j, part in enumerate(parts[2:], 2):
                            if part.startswith("$"):
                                price_start_idx = j
                                break
                            size_parts.append(part)

                        size = " ".join(size_parts) if size_parts else ""

                        retail_us, retail_ca, multi_price = "", "", ""
                        sku, article, quantity = "", "", ""

                        if price_start_idx is not None and price_start_idx + 1 < len(parts):
                            retail_us = parts[price_start_idx]
                            retail_ca = parts[price_start_idx + 1]

                            for k in range(price_start_idx + 2, len(parts)):
                                if len(parts[k]) == 13 and parts[k].isdigit():
                                    sku = parts[k]
                                elif len(parts[k]) == 8 and parts[k].isdigit():
                                    article = parts[k]
                                elif parts[k].isdigit():
                                    quantity = parts[k]

                        # ✅ Skip empty rows
                        if any([style, color_code, size, sku, article, quantity]):
                            table_rows.append({
                                "style": style,
                                "color_code": color_code,
                                "size": size,
                                "sku_desc": sku,  # ✅ renamed
                                "panty_length_2": "",
                                "retail_us": retail_us,
                                "retail_ca": retail_ca,
                                "multi_price": multi_price,
                                "article": article,
                                "quantity": quantity
                            })

                elif table_format == 'basic':
                    parts = re.split(r'\s+', line)
                    if len(parts) >= 6:
                        style = parts[0]
                        color_code = parts[1]
                        size = parts[2]

                        sku, article, quantity = "", "", ""

                        for j, part in enumerate(parts[3:], 3):
                            if len(part) == 13 and part.isdigit():
                                sku = part
                                if j + 1 < len(parts) and len(parts[j+1]) == 8 and parts[j+1].isdigit():
                                    article = parts[j+1]
                                    if j + 2 < len(parts) and parts[j+2].isdigit():
                                        quantity = parts[j+2]
                                break

                        if not quantity and parts[-1].isdigit():
                            quantity = parts[-1]

                        # ✅ Skip empty rows
                        if any([style, color_code, size, sku, article, quantity]):
                            table_rows.append({
                                "style": style,
                                "color_code": color_code,
                                "size": size,
                                "sku_desc": sku,  # ✅ renamed
                                "panty_length_2": "",
                                "retail_us": "",
                                "retail_ca": "",
                                "multi_price": "",
                                "article": article,
                                "quantity": quantity
                            })

        # Process extracted rows
        for row in table_rows:
            extracted_data["PO Number"].append(po_number)
            extracted_data["Item Code"].append(item_code)
            extracted_data["Product Code"].append(product_code)
            extracted_data["Style"].append(row["style"])
            extracted_data["Color Code"].append(row["color_code"])
            extracted_data["SO Number"].append(so_number)
            extracted_data["Line Number"].append(line_number)
            extracted_data["Size"].append(row["size"])
            extracted_data["SKU Desc"].append(row["sku_desc"])  # ✅ updated
            extracted_data["Panty Length 2"].append(row["panty_length_2"])
            extracted_data["Retail (US)"].append(row.get("retail_us", ""))
            extracted_data["Retail (CA)"].append(row.get("retail_ca", ""))
            extracted_data["Multi Price"].append(row.get("multi_price", ""))
            extracted_data["Product Desc"].append(product_desc)
            extracted_data["Article"].append(row["article"])
            extracted_data["Quantity"].append(row["quantity"])

    except Exception as e:
        st.error(f"Error extracting data from {uploaded_file.name}: {str(e)}")
        import traceback
        st.error(f"Traceback: {traceback.format_exc()}")

    return pd.DataFrame(extracted_data)

# ---------------------- Main UI ----------------------
st.set_page_config(
    page_title="MAS Data Entry Checking Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for beautiful styling
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
        padding: 2rem;
        border-radius: 15px;
        text-align: center;
        color: white;
        margin-bottom: 2rem;
    }
    
    .upload-section {
        background: #f8f9fa;
        padding: 2rem;
        border-radius: 15px;
        border: 2px dashed #007bff;
        margin-bottom: 1rem;
    }
    
    .results-section {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        border-left: 5px solid #28a745;
    }
    
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin: 0.5rem;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding-left: 20px;
        padding-right: 20px;
        background-color: #f0f2f6;
        border-radius: 10px;
        border: none;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# Main Header
st.markdown("""
<div class="main-header">
    <h1>📄 MAS Data Entry Checking Tool – Price Tickets</h1>
    <p>Upload PDF files to extract PO numbers, product codes, and detailed data</p>
</div>
""", unsafe_allow_html=True)

# ---------------------- Sidebar File Upload ----------------------
with st.sidebar:
    st.markdown("## 📁 File Upload")
    
    # Price Tickets Upload Section
    st.markdown("### 🎟️ Perches Order PDFs")
    uploaded_tickets = st.file_uploader(
        "Choose Price Tickets PDF files", 
        type=["pdf"], 
        accept_multiple_files=True,
        key="price_tickets",
        help="Upload one or more Price Tickets PDF files"
    )
    
    if uploaded_tickets:
        st.success(f"✅ {len(uploaded_tickets)} file(s) uploaded!")
        with st.expander("📄 Uploaded Files", expanded=False):
            for file in uploaded_tickets:
                st.write(f"• {file.name}")
    
    st.markdown("---")
    
    # Work Order Upload Section
    st.markdown("### 📋 Work Order PDFs")
    uploaded_wo = st.file_uploader(
        "Choose WO PDF files", 
        type=["pdf"], 
        accept_multiple_files=True,
        key="work_orders",
        help="Upload one or more Work Order PDF files"
    )
    
    if uploaded_wo:
        st.success(f"✅ {len(uploaded_wo)} file(s) uploaded!")
        with st.expander("📄 Uploaded Files", expanded=False):
            for file in uploaded_wo:
                st.write(f"• {file.name}")

    st.markdown("---")
    st.markdown("### 🔧 Support")
    st.markdown("For technical support or questions, please contact the Razz.. development team.")

# Create two tabs for different PDF types
tab1, tab2 = st.tabs(["🎟️ Perches Order Results", "📋 Work Order Results"])


# ---------------------- Price Tickets Tab ----------------------
with tab1:
    if uploaded_tickets:
        st.markdown('<div class="results-section">', unsafe_allow_html=True)
        st.subheader("📊 Extraction Results - Price Tickets")
        
        rows = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, up in enumerate(uploaded_tickets):
            status_text.text(f"Processing {up.name}...")
            progress_bar.progress(int(((idx) / len(uploaded_tickets)) * 100))
            try:
                raw = up.read()
                pages = read_pdf_text(raw)
                text = full_text(pages)
                
                po_number = extract_po_number(text)
                product_codes = extract_product_codes(text)
                
                # Process each product code
                if product_codes:
                    for code_data in product_codes:
                        rows.append(
                            {
                                "PO Number": po_number,
                                "Item Code": code_data["Item Code"],
                                "Product Code": code_data["Product Code"],
                                "Style": code_data["Style"],
                                "Color Code": code_data["Color Code"],
                                "SO Number": code_data["SO Number"],
                                "Line Number": code_data["Line Number"],
                                "Size": code_data["Size"],
                                "Size Qty": code_data["Size Qty"]
                            }
                        )
                else:
                    # If no product codes found, add a row with PO number only
                    rows.append(
                        {
                            "PO Number": po_number,
                            "Item Code": None,
                            "Product Code": None,
                            "Style": None,
                            "Color Code": None,
                            "SO Number": None,
                            "Line Number": None,
                            "Size": None,
                            "Size Qty": None
                        }
                    )
            except Exception as e:
                st.error(f"Error processing {up.name}: {str(e)}")
                continue
                
        progress_bar.progress(100)
        status_text.text("✅ Processing complete!")
        
        if rows:
            # Process rows to avoid repeating SO numbers
            processed_rows = []
            last_so = None
            
            for row in rows:
                # Skip rows with no SO Number
                if row["SO Number"] is None:
                    processed_rows.append(row)
                    continue
                    
                # If this SO Number is different from the last one, keep it
                if row["SO Number"] != last_so:
                    processed_rows.append(row)
                    last_so = row["SO Number"]
                else:
                    # Create a copy with blank SO Number
                    new_row = row.copy()
                    new_row["SO Number"] = ""
                    processed_rows.append(new_row)
            
            df = pd.DataFrame(processed_rows)
            display_columns = [
                "PO Number",
                "Item Code",
                "Product Code",
                "Style",
                "Color Code",
                "SO Number",
                "Line Number",
                "Size",
                "Size Qty"
            ]
            
            # Display metrics
            col1_m, col2_m, col3_m = st.columns(3)
            with col1_m:
                st.metric("📄 Total Files", len(uploaded_tickets))
            with col2_m:
                st.metric("📊 Total Records", len(df))
            with col3_m:
                unique_pos = df["PO Number"].nunique()
                st.metric("📋 Unique POs", unique_pos)
            
            st.dataframe(df[display_columns], use_container_width=True)
            
            # Download button
            csv_data = df[display_columns].to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download Price Tickets CSV",
                data=csv_data,
                file_name="price_tickets_extracted.csv",
                mime="text/csv",
                use_container_width=True
            )
            
            # Add a section to show all product codes for each PO
            with st.expander("📋 Detailed View by PO Number", expanded=False):
                po_groups = df.groupby("PO Number")
                for po, group in po_groups:
                    st.write(f"**PO Number: {po}**")
                    codes_data = group[["Item Code", "Product Code", "Style", "Color Code", "SO Number", 
                                       "Line Number", "Size", "Size Qty"]].copy()
                    st.dataframe(codes_data, use_container_width=True)
                    
                    # Show SO numbers for this PO
                    st.write(f"**SO Numbers for PO {po}:**")
                    so_numbers = group["SO Number"].unique()
                    for so in so_numbers:
                        if so:
                            st.write(f"• {so}")
        else:
            st.warning("⚠️ No data could be extracted from the uploaded files.")
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("👈 Upload Price Tickets PDF files in the sidebar to see results here")

# ---------------------- Work Order Tab ----------------------
with tab2:
    if uploaded_wo:
        st.markdown('<div class="results-section">', unsafe_allow_html=True)
        st.subheader("📊 Extraction Results - Work Orders")
        
        all_dfs = []
        individual_summaries = []
        
        with st.spinner("Processing WO PDF(s) and extracting data..."):
            for file in uploaded_wo:
                df = extract_data_from_pdf(file)
                if not df.empty:
                    df["Source File"] = file.name  # keep track of source
                    all_dfs.append(df)
                    
                    # Store individual summary data
                    individual_summaries.append({
                        "File Name": file.name,
                        "PO Number": df["PO Number"].iloc[0] if not df.empty else "",
                        "Total Items": len(df),
                        "Total Quantity": pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int).sum(),
                        "Unique Styles": df["Style"].nunique()
                    })

        if all_dfs:
            # Display overall metrics
            col1_m, col2_m, col3_m, col4_m = st.columns(4)
            with col1_m:
                st.metric("📄 Total Files", len(all_dfs))
            with col2_m:
                total_items = sum(len(df) for df in all_dfs)
                st.metric("📊 Total Items", total_items)
            with col3_m:
                combined_df = pd.concat(all_dfs, ignore_index=True)
                total_qty = pd.to_numeric(combined_df["Quantity"], errors="coerce").fillna(0).astype(int).sum()
                st.metric("📦 Total Quantity", f"{total_qty:,}")
            with col4_m:
                st.metric("🎨 Unique Styles", combined_df["Style"].nunique())
            
            # Display individual tables for each WO
            st.write("### 📋 Individual WO Extractions")
            
            for i, df in enumerate(all_dfs):
                file_name = df["Source File"].iloc[0]
                po_number = df["PO Number"].iloc[0] if not df.empty else "Unknown"
                
                with st.expander(f"📄 {file_name} (PO: {po_number})", expanded=True):
                    st.dataframe(df.drop(columns=["Source File"]), use_container_width=True)
                    
                    # Individual file summary
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Items", len(df))
                    with col2:
                        file_total_qty = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int).sum()
                        st.metric("Total Quantity", f"{file_total_qty:,}")
                    with col3:
                        st.metric("Unique Styles", df["Style"].nunique())
                    
                    # Individual download button
                    csv_individual = df.drop(columns=["Source File"]).to_csv(index=False).encode("utf-8")
                    st.download_button(
                        label=f"📥 Download {file_name} CSV",
                        data=csv_individual,
                        file_name=f"{file_name.replace('.pdf', '')}_extracted.csv",
                        mime="text/csv",
                        key=f"download_wo_{i}",
                        use_container_width=True
                    )

            # Combined data section
            st.write("### 📊 Combined Summary")
            
            # Summary table of all files
            summary_df = pd.DataFrame(individual_summaries)
            st.dataframe(summary_df, use_container_width=True)

            # Combined download option
            csv_combined = combined_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download All WO Combined CSV",
                data=csv_combined,
                file_name="all_wo_combined_data.csv",
                mime="text/csv",
                use_container_width=True
            )
            
            # Optional: Show combined data table in expander
            with st.expander("📋 View Combined Data Table", expanded=False):
                st.dataframe(combined_df.drop(columns=["Source File"]), use_container_width=True)
        else:
            st.warning("⚠️ No data extracted from uploaded WO PDFs.")
        
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("👈 Upload Work Order PDF files in the sidebar to see results here")


# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 1rem;">
     🚀 <strong>Customer Care System v2.0</strong> | 
        Powered by <strong>Razz....</strong> | 
        Advanced PDF Analysis & Comparison Technology
</div>
""", unsafe_allow_html=True)