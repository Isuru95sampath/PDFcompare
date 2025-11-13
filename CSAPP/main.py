import streamlit as st
import pandas as pd
import re
import pdfplumber  # Add this import
from io import BytesIO

# Import all the modules
from ui_config import configure_page, apply_custom_css, display_header, display_footer
from auth import setup_sidebar
from logging_utils import log_to_text
from excel_utils import read_excel_table, process_excel_table_data
from pdf_utils import (
    uploaded_file_to_bytesio, 
    create_styles_pdf, 
    merge_pdfs_with_po,
    extract_style_numbers_from_po_first_page, 
    extract_po_number, 
    extract_so_number_from_wo,
    extract_all_so_numbers_from_wo, 
    extract_wo_fields, 
    extract_po_fields,
    extract_wo_items_table, 
    extract_po_details, 
    reorder_wo_by_size, 
    reorder_po_by_size,
    debug_po_extraction, 
    compare_addresses, 
    check_vsba_in_po_line,
    extract_item_description_product_code_and_check_vsba,
    extract_wo_product_code_with_vsba, 
    extract_po_product_code_with_vsba,
    compare_vsba_status
)
from data_comparison import (
    enhanced_quantity_matching, 
    compare_codes, 
    get_excel_style_number, 
    update_po_details_with_excel_styles,
    update_matched_items_with_excel_styles, 
    combine_wo_and_excel_data,
    update_so_color_display, 
    clean_product_code,
    fill_empty_style_2_from_excel  
)

def show_progress_steps(current_step=1):
    return ""

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

def show_wo_po_summary(wo_df_detailed, po_details, metrics):
    """
    Display main summary view with metrics, WO and PO details.
    """

    # ------------------- Metrics Section -------------------
    st.markdown("""
    <div class="section-header">
        <h2 class="section-title">📈 WO vs PO Summary</h2>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    col1.metric("✅ Matches", metrics.get('matches', 0))
    col2.metric("❌ Mismatches", metrics.get('mismatches', 0))
    col3.metric("⚠️ Missing Items", metrics.get('missing', 0))

    st.markdown("<hr>", unsafe_allow_html=True)

    # ------------------- Tables Section -------------------
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 🧾 Work Order (WO) Items")
        if not wo_df_detailed.empty:
            # Drop empty columns dynamically
            for col in wo_df_detailed.columns:
                if wo_df_detailed[col].isnull().all() or (wo_df_detailed[col].astype(str).str.strip() == '').all():
                    wo_df_detailed = wo_df_detailed.drop(columns=[col])
            if 'WO Product Code' in wo_df_detailed.columns:
                wo_df_detailed = wo_df_detailed.drop(columns=['WO Product Code'])

            st.dataframe(wo_df_detailed, use_container_width=True, hide_index=True)

    with col2:
        st.markdown("### 📋 Purchase Order (PO) Items")
        po_df = pd.DataFrame(po_details)
        st.dataframe(po_df, use_container_width=True, hide_index=True)

    # ------------------- SO vs WO Color Comparison -------------------
    st.markdown("<hr>", unsafe_allow_html=True)

    example_so_numbers = ["RED-01", "BLUE-02", "RED-03"]
    example_wo_items = [
        {"WO Colour Code": "RED"},
        {"WO Colour Code": "BLUE"},
        {"WO Colour Code": "RED"},
        {"WO Colour Code": "GREEN"},
        {"WO Colour Code": "RED"},
    ]

    update_so_color_display(example_so_numbers, example_wo_items)

def main():
    # Initialize session state if it doesn't exist
    if 'processed_excel_data' not in st.session_state:
        st.session_state.processed_excel_data = None
    
    # Configure the page
    configure_page()
    apply_custom_css()
    display_header()
    
    # Setup sidebar and get user selection and uploaded files
    selected_user, wo_file, po_file = setup_sidebar()
    
    # -------------------- Excel/PDF Merger Section --------------------
    with st.expander("📓 Excel Table Data Extractor", expanded=False):
        st.markdown("""
        <div class="section-header">
            <h3 class="section-title">📊 Excel Table Data Extractor</h3>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        # In main.py, replace the excel_file uploader section with this:

        with col1:
            excel_files = st.file_uploader(
                "📤 Upload Excel Files (Multiple)", 
                type=["xls", "xlsx"], 
                key="excel_merger",
                disabled=not selected_user,
                accept_multiple_files=True,
                help="Upload one or more Excel files containing table data starting at row 22"
            )
        
        with col2:
            pdf_file_merger = st.file_uploader(
                "📄 Upload PDF File (Optional)", 
                type=["pdf"], 
                key="pdf_merger",
                disabled=not selected_user,
                help="Optional: Upload PDF file to merge with extracted styles"
            )
        
        if excel_files:
            with st.spinner("🔄 Extracting table data from multiple files..."):
                # Process each Excel file
                all_files_data = []
                all_styles = []
                
                for excel_file in excel_files:
                    # Extract table data from all sheets of this file
                    all_sheets_data = read_excel_table(excel_file)
                    
                    # Add file information to each sheet
                    for sheet_data in all_sheets_data:
                        sheet_data['file_name'] = excel_file.name
                    
                    all_files_data.extend(all_sheets_data)
                    
                    # Extract styles from this file
                    sheets_with_data = [s for s in all_sheets_data if not s['data'].empty]
                    for sheet_data in sheets_with_data:
                        if sheet_data['style_number']:
                            all_styles.append(sheet_data['style_number'])
                
                # Filter out sheets with no data
                sheets_with_data = [s for s in all_files_data if not s['data'].empty]
                
                if sheets_with_data:
                    processed_data = process_excel_table_data(sheets_with_data)
                    
                    # Store processed data in session state
                    st.session_state.processed_excel_data = processed_data
                    
                    # If no style numbers found in STYLE columns, try to extract from the data
                    if not all_styles:
                        for sheet_data in sheets_with_data:
                            df = sheet_data['data']
                            if not df.empty:
                                # Look for any 8-digit number in the dataframe
                                for col in df.columns:
                                    for val in df[col].dropna():
                                        val_str = str(val).strip()
                                        if re.match(r'^\d{8}$', val_str):
                                            all_styles.append(val_str)
                                            break
                                    if all_styles:
                                        break
                                if all_styles:
                                    break
                    
                    # Display the extracted table data
                    st.markdown(f"""
                    <div class="alert-success">
                        ✅ <strong>Success!</strong> Table data extracted from {len(excel_files)} file(s) successfully.
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Group sheets by file
                    files_dict = {}
                    for sheet_data in all_files_data:
                        file_name = sheet_data['file_name']
                        if file_name not in files_dict:
                            files_dict[file_name] = []
                        files_dict[file_name].append(sheet_data)
                    
                    # Display each file's sheets
                    for file_name, sheets in files_dict.items():
                        with st.expander(f"📁 File: {file_name}"):
                            for sheet_data in sheets:
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
                        if all_styles and pdf_file_merger:
                            styles_pdf = create_styles_pdf(all_styles)
                            
                            # Check if PO file is available in session state
                            po_pdf_for_merge = None
                            if 'po_file' in st.session_state and st.session_state.po_file is not None:
                                po_pdf_for_merge = uploaded_file_to_bytesio(st.session_state.po_file)
                            
                            # Convert the uploaded PDF files to BytesIO objects
                            pdf_merger_bytes = uploaded_file_to_bytesio(pdf_file_merger)
                            
                            # Merge PDFs with PO if available
                            final_pdf = merge_pdfs_with_po(styles_pdf, pdf_merger_bytes, po_pdf_for_merge)
                            
                            if final_pdf:
                                if po_pdf_for_merge:
                                    st.download_button(
                                        label="⬇️ Download Merged PDF (with PO)",
                                        data=final_pdf,
                                        file_name="Merged-PO-WO.pdf",
                                        mime="application/pdf",
                                        use_container_width=True
                                    )
                                else:
                                    st.download_button(
                                        label="⬇️ Download Merged PDF",
                                        data=final_pdf,
                                        file_name="Merged-PO.pdf",
                                        mime="application/pdf",
                                        use_container_width=True
                                    )
                        elif all_styles and not pdf_file_merger:
                            st.info("Upload a PDF file to merge with extracted styles")
                else:
                    st.markdown("""
                    <div class="alert-warning">
                        ⚠️ <strong>Warning:</strong> No valid table data found in any sheet.
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="alert-info">
                ℹ️ <strong>Info:</strong> Please upload one or more Excel files to extract table data.
            </div>
            """, unsafe_allow_html=True)

    # -------------------- Main Analysis Section --------------------
    if selected_user and wo_file and po_file:
        with st.spinner("🔄 Processing files and analyzing data..."):
            wo = extract_wo_fields(wo_file)
            po = extract_po_fields(po_file)
            wo_items = extract_wo_items_table(wo_file, wo["product_codes"])
            wo_items = reorder_wo_by_size(wo_items)
            
            # Updated to handle new return format with PO product codes from Item column
            po_details_result = extract_po_details(po_file)
            po_details_raw = po_details_result["po_items"]
            po_product_codes_from_item = po_details_result.get("po_product_codes_from_item", [])
            po_details = reorder_po_by_size(po_details_raw)
            
            addr_res = compare_addresses(wo, po)
            
            # Updated to include PO product codes from Item column
            code_res = compare_codes(po_details, wo_items, po_product_codes_from_item)
            code_table_df = compare_codes(po_details, wo_items, po_product_codes_from_item)
            
            matched, mismatched = enhanced_quantity_matching(wo_items, po_details)
            po_number = extract_po_number(po_file)
            so_numbers = extract_all_so_numbers_from_wo(wo_file)
            
            # NEW: Check if VSBA is in the same line as PO number
            vsba_in_po_line = check_vsba_in_po_line(po_file)
            
            # NEW: Extract product code from Item Description and check for VSBA
            item_desc_product_code, vsba_in_item_desc = extract_item_description_product_code_and_check_vsba(po_file)
            
            # NEW: Update Style 2 from Excel if missing from PO
            # Fill empty Style 2 values from Excel data if available
            if hasattr(st.session_state, 'processed_excel_data') and st.session_state.processed_excel_data is not None:
                matched, mismatched = fill_empty_style_2_from_excel(
                    matched, mismatched, st.session_state.processed_excel_data
                )

            # Get Excel style number (if available)
            excel_style_number = None
            if hasattr(st.session_state, 'processed_excel_data') and st.session_state.processed_excel_data is not None:
                excel_style_number = get_excel_style_number(st.session_state.processed_excel_data)

            # ...

            # Update the call to enhanced_quantity_matching to include the excel_style_number
            matched, mismatched = enhanced_quantity_matching(wo_items, po_details, tolerance=0, excel_style=excel_style_number)
            
            # Check if PO has style numbers
            po_has_styles = any(po_item.get("Style 2", "") for po_item in po_details)
            
            # DO NOT merge Excel styles into PO details anymore
            # The style number will be displayed separately at the top
            
            # Create code_table_df here to ensure it's available in the scope
            # Include PO product codes from Item column in the comparison
            po_all_codes = [po.get("Product_Code", "").strip().upper() for po in po_details if po.get("Product_Code")]

            # Add PO product codes from Item column
            if po_product_codes_from_item:
                po_all_codes.extend([code.strip().upper() for code in po_product_codes_from_item])

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
        
        st.markdown(show_progress_steps(4), unsafe_allow_html=True)

        st.markdown("""
        <div class="alert-success">
            🎉 <strong>Analysis Complete!</strong> Your files have been processed successfully.
        </div>
        """, unsafe_allow_html=True)

        # Display Excel style number at the very top (only once)
        if excel_style_number:
            st.markdown(f"""
            <div class="alert-info" style="text-align: center; font-size: 1.1rem;">
                <strong>Style Number:</strong> {excel_style_number}
            </div>
            """, unsafe_allow_html=True)

        #Analysis Overview Section
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
        
        # Display the code_table_df that was created earlier
        st.dataframe(code_table_df, use_container_width=True, hide_index=True)


         # Extract VSBA information from WO
        wo_vsba_data = extract_wo_product_code_with_vsba(wo_file)
        
        # Extract VSBA information from PO
        po_vsba_data = extract_po_product_code_with_vsba(po_file)
        
        # Compare VSBA status
        vsba_comparison = compare_vsba_status(wo_vsba_data, po_vsba_data)
        
        # Display WO Product Codes with VSBA Status (only first row)
        st.markdown("#### 📄 Work Order (WO) Product Codes")
        if wo_vsba_data:
            # Show only the first row
            wo_vsba_df = pd.DataFrame([wo_vsba_data[0]])
            st.dataframe(wo_vsba_df, use_container_width=True, hide_index=True)
        else:
            st.info("No product codes found in WO")
        
        # Display PO Product Codes with VSBA Status (only rows where VSBA is found)
        st.markdown("#### 📋 Purchase Order (PO) Product Codes")
        if po_vsba_data:
            # Filter to show only rows where VSBA is found
            vsba_found_rows = [item for item in po_vsba_data if item["Has_VSBA"]]
            
            if vsba_found_rows:
                po_vsba_df = pd.DataFrame(vsba_found_rows)
                st.dataframe(po_vsba_df, use_container_width=True, hide_index=True)
            else:
                st.info("No product codes with VSBA found in PO")
        else:
            st.info("No product codes found in PO")
        
        # Display VSBA Comparison Summary
        st.markdown("#### 📊 VSBA Comparison Summary")
        vsba_summary_df = pd.DataFrame([{
            "WO has VSBA": "✅ Yes" if vsba_comparison["WO_VSBA_Found"] else "❌ No",
            "PO has VSBA": "✅ Yes" if vsba_comparison["PO_VSBA_Found"] else "❌ No",
            "Overall Status": vsba_comparison["Status"]
        }])
        st.dataframe(vsba_summary_df, use_container_width=True, hide_index=True)
        
        # Display alert based on VSBA status
        if vsba_comparison["Both_Have_VSBA"]:
            st.markdown("""
            <div class="alert-success">
                ✅ <strong>VSBA Match:</strong> Both WO and PO contain VSBA in their product codes.
            </div>
            """, unsafe_allow_html=True)
        elif vsba_comparison["WO_VSBA_Found"] or vsba_comparison["PO_VSBA_Found"]:
            st.markdown("""
            <div class="alert-warning">
                ⚠️ <strong>VSBA Mismatch:</strong> VSBA found in only one document. Please verify.
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="alert-info">
                ℹ️ <strong>No VSBA:</strong> VSBA not found in either WO or PO product codes.
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<hr>", unsafe_allow_html=True)
        
        st.markdown("""
        <div class="section-header">
            <h3 class="section-title">📋 Matching other items</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # Check if we need to update PO details with style numbers from processed Excel data
        # First, check if any Style 2 values are empty in the matched and mismatched data
        has_empty_style_2 = False
        for item in matched + mismatched:
            if not item.get("Style 2", ""):
                has_empty_style_2 = True
                break
        
        # If there are empty Style 2 values and we have processed Excel data, update them
        if has_empty_style_2 and hasattr(st.session_state, 'processed_excel_data') and st.session_state.processed_excel_data is not None:
            # Update matched and mismatched data with style numbers from processed Excel data
            updated_matched = []
            for match_item in matched:
                # Only update if Style 2 is empty
                if not match_item.get("Style 2", ""):
                    # Try to find a matching style from the processed Excel data
                    wo_style = match_item.get("Style", "")
                    wo_color = match_item.get("WO Colour Code", "")
                    wo_size = match_item.get("WO Size", "")
                    
                    # Look for matching style in processed Excel data
                    excel_style = None
                    if hasattr(st.session_state, 'processed_excel_data') and st.session_state.processed_excel_data is not None:
                        excel_data = st.session_state.processed_excel_data
                        for _, row in excel_data.iterrows():
                            excel_style_val = str(row.get('Excel Style', '')).strip()
                            excel_color_val = str(row.get('Excel Colour Code', '')).strip().upper()
                            excel_size_val = str(row.get('Excel Size', '')).strip().upper()
                            
                            if (excel_style_val and 
                                excel_color_val == wo_color.upper() and 
                                excel_size_val == wo_size.upper()):
                                excel_style = excel_style_val
                                break
                    
                    # Update the match item with the found style
                    updated_match_item = match_item.copy()
                    if excel_style:
                        updated_match_item["Style 2"] = excel_style
                    
                    updated_matched.append(updated_match_item)
                else:
                    # Keep the original item if Style 2 is not empty
                    updated_matched.append(match_item)

              
                if hasattr(st.session_state, 'processed_excel_data') and st.session_state.processed_excel_data is not None:
                    # Check if PO has Style 2
                    po_has_style_2 = any(po_item.get("Style 2", "") for po_item in po_details)
                    
                    # Check if matched or mismatched items have Style 2
                    items_have_style_2 = any(item.get("Style 2", "") for item in matched + mismatched)
                    
                    # If PO doesn't have Style 2 but items do, it means Excel style was used
                    if not po_has_style_2 and items_have_style_2:
                        excel_style = None
                        for _, row in st.session_state.processed_excel_data.iterrows():
                            style = str(row.get('Style', '')).strip()
                            if style:
                                excel_style = style
                                break
                        
                        if excel_style:
                            st.markdown(f"""
                            <div class="alert-info">
                                ℹ️ <strong>Style Number from Excel:</strong> {excel_style} (used because PO was missing style information)
                            </div>
                            """, unsafe_allow_html=True)
            
            # Update mismatched data with style numbers from processed Excel data
            updated_mismatched = []
            for mismatch_item in mismatched:
                # For items with WO data and empty Style 2
                if mismatch_item.get("Style") and not mismatch_item.get("Style 2", ""):
                    wo_style = mismatch_item.get("Style", "")
                    wo_color = mismatch_item.get("WO Colour Code", "")
                    wo_size = mismatch_item.get("WO Size", "")
                    
                    # Look for matching style in processed Excel data
                    excel_style = None
                    if hasattr(st.session_state, 'processed_excel_data') and st.session_state.processed_excel_data is not None:
                        excel_data = st.session_state.processed_excel_data
                        for _, row in excel_data.iterrows():
                            excel_style_val = str(row.get('Excel Style', '')).strip()
                            excel_color_val = str(row.get('Excel Colour Code', '')).strip().upper()
                            excel_size_val = str(row.get('Excel Size', '')).strip().upper()
                            
                            if (excel_style_val and 
                                excel_color_val == wo_color.upper() and 
                                excel_size_val == wo_size.upper()):
                                excel_style = excel_style_val
                                break
                    
                    # Update the mismatch item with the found style
                    updated_mismatch_item = mismatch_item.copy()
                    if excel_style:
                        updated_mismatch_item["Style 2"] = excel_style
                    
                    updated_mismatched.append(updated_mismatch_item)
                else:
                    # For items with PO data only or already filled Style 2, keep as is
                    updated_mismatched.append(mismatch_item)
            
            # Use the updated matched and mismatched data
            matched = updated_matched
            mismatched = updated_mismatched
        
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
            <h3 class="section-title">📊 Combined WO and Excel Data </h3>
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
        
        # Initialize combined_df as an empty DataFrame
        combined_df = pd.DataFrame()
        
        # Check if we have Excel data
        if hasattr(st.session_state, 'processed_excel_data') and st.session_state.processed_excel_data is not None:
            excel_df = st.session_state.processed_excel_data
            
            if not excel_df.empty:
                # Combine the data using the enhanced function
                with st.spinner("🔄 Combining WO and Excel data..."):
                    combined_df = combine_wo_and_excel_data(wo_df, excel_df)
                
                # Now combined_df is guaranteed to be a DataFrame (not None)
                if not combined_df.empty:
                    # Count matches and mismatches
                    total_rows = len(combined_df)
                    full_matches = len(combined_df[combined_df["Overall Match"] == "✅ Full Match"])
                    mismatches = len(combined_df[combined_df["Overall Match"] == "❌ Mismatch"])
                    missing_wo = len(combined_df[combined_df["Overall Match"] == "⚠️ Missing WO Data"])
                    missing_excel = len(combined_df[combined_df["Overall Match"] == "⚠️ Missing Excel Data"])
                    
                    # Display match statistics
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-value">{total_rows}</div>
                            <div class="metric-label">Total Rows</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-card" style="border-top-color: #28a745;">
                            <div class="metric-value" style="color: #28a745;">{full_matches}</div>
                            <div class="metric-label">Full Matches</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f"""
                        <div class="metric-card" style="border-top-color: #dc3545;">
                            <div class="metric-value" style="color: #dc3545;">{mismatches}</div>
                            <div class="metric-label">Mismatches</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col4:
                        st.markdown(f"""
                        <div class="metric-card" style="border-top-color: #ffc107;">
                            <div class="metric-value" style="color: #e67e22;">{missing_wo + missing_excel}</div>
                            <div class="metric-label">Missing Data</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Display the combined table
                    st.markdown("### 📊 Combined WO and Excel Data")
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
                        ⚠️ <strong>Empty Combined Data:</strong> The combined data table is empty. This might be due to no matching rows between WO and Excel data.
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div class="alert-warning">
                    ⚠️ <strong>Empty Excel Data:</strong> The processed Excel data is empty.
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
        
                # SO Number and WO Color Code section
        so_color_df = update_so_color_display(so_numbers, wo_items)
        
        # Check for perfect matches in both tables
        combined_perfect_match = False
        so_color_perfect_match = False
        
        # Check Combined WO and Excel Data table
        if not combined_df.empty:
            if "Overall Match" in combined_df.columns:
                # Check if all rows have "✅ Full Match"
                combined_perfect_match = all(combined_df["Overall Match"] == "✅ Full Match")
        
        # Check WO Color Codes vs SO Numbers Comparison table
        if not so_color_df.empty and "Status" in so_color_df.columns:
            # Check if all rows have "✅ Match"
            so_color_perfect_match = all(so_color_df["Status"] == "✅ Match")
        
        address_ok = addr_res.get("Status", "") == "✅ Match"
        codes_ok = not code_table_df.empty and all(code_table_df["🔍 Match Status"].isin(["✅ Exact Match", "✅ Partial Match"]))
        matched_df = pd.DataFrame(matched) if matched else pd.DataFrame()
        matched_ok = not matched_df.empty and all(matched_df["Status"] == "🟩 Full Match")
        mismatched_empty = len(mismatched) == 0

        # Check VSBA status for balloon condition - more robust handling
        vsba_ok = False
        vsba_status = None
        
        # Handle both dictionary and DataFrame cases for vsba_comparison
        if isinstance(vsba_comparison, dict):
            vsba_status = vsba_comparison.get("Status", "")
        elif isinstance(vsba_comparison, pd.DataFrame) and not vsba_comparison.empty:
            if "Status" in vsba_comparison.columns:
                # Get the first row's status or check if all rows have the same status
                vsba_status = vsba_comparison["Status"].iloc[0]
            elif "Overall Status" in vsba_comparison.columns:
                vsba_status = vsba_comparison["Overall Status"].iloc[0]
        
        # Check VSBA conditions
        if vsba_status == "✅ Both have VSBA":
            vsba_ok = True
        elif vsba_status == "❌ Neither has VSBA":
            vsba_ok = True
        else:
            vsba_ok = False
            
        # Debug: Uncomment these lines to see what's happening
        # st.write(f"VSBA Status: {vsba_status}")
        # st.write(f"VSBA OK: {vsba_ok}")

        # Updated condition: All checks must pass AND VSBA condition must be satisfied
        if address_ok and codes_ok and matched_ok and mismatched_empty and combined_perfect_match and so_color_perfect_match and vsba_ok:
            match_status = "PERFECT MATCH!"
            st.markdown("""
            <audio autoplay>
                <source src="">
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
            
            # Create a detailed mismatch message
            issues = []
            if not address_ok:
                issues.append("Address mismatch")
            if not codes_ok:
                issues.append("Product code mismatch")
            if not matched_ok or not mismatched_empty:
                issues.append("Item matching issues")
            if not combined_perfect_match:
                issues.append("WO/Excel data mismatch")
            if not so_color_perfect_match:
                issues.append("SO/Color mismatch")
            if not vsba_ok:
                if vsba_status:
                    issues.append(f"VSBA mismatch (Status: {vsba_status})")
                else:
                    issues.append("VSBA mismatch (status not found)")
            
            issues_text = ", ".join(issues) if issues else "Some data points need verification"
            
            st.markdown(f"""
            <div class="alert-warning">
                ⚠️ <strong>Review Required:</strong> {issues_text}. Check the details below.
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
            so_numbers_str = "; ".join(so_numbers) if so_numbers else ""
            success, message = log_to_text(
                selected_user, 
                first_product_code, 
                first_reference, 
                match_status, 
                po_number,
                so_numbers_str
            )
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
    
    # Display the footer
    display_footer()

if __name__ == "__main__":
    main()