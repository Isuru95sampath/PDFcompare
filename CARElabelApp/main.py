import streamlit as st
import os
import tempfile
import shutil
import sys
import pandas as pd
import re
import pdfplumber

# Import modules
from ui_components import initialize_page, initialize_session_state, create_sidebar, display_wo_details
from email_processor import process_email_to_pdf
from wo_extractor import process_wo_file, extract_wo_items_table_enhanced, extract_size_breakdown_table_robust, extract_and_sort_wo_sizes
# Corrected and Consolidated po_extractor imports
from po_extractor import (
    display_email_po_debug_info, 
    extract_garment_description_table, 
    extract_merged_po_details, 
    display_merged_po_results, 
    extract_po_numbers_from_email_body, 
    extract_email_body_data, 
    filter_garment_description_by_po,
    # THIS IS THE CORRECTED NAME THAT FIXES THE IMPORTERROR
    extract_email_body_item_data 
)
from comparison import display_comparison_table as display_detailed_comparison_table, display_size_comparison_for_matched_po


def extract_po_size_breakdown(po_list):
    """
    Extract size and quantity breakdown from PO data
    Returns a dictionary with sizes as keys and quantities as values
    """
    po_sizes = {}
    
    # Check if po_list is valid
    if not isinstance(po_list, list):
        st.error("PO data is not in the expected format (list). Cannot extract size breakdown.")
        return {}

    for po in po_list:
        # Check if po is a dictionary and has 'items'
        if not isinstance(po, dict) or 'items' not in po or not isinstance(po['items'], list):
            continue
            
        for item in po['items']:
            # Check if item is a dictionary
            if not isinstance(item, dict):
                continue
                
            size = item.get('size', '').strip()
            if not size:
                continue
                
            quantity = 0
            if item.get('quantity') and item['quantity'] != '':
                try:
                    quantity = float(item['quantity'])
                except (ValueError, TypeError):
                    quantity = 0
            
            if size in po_sizes:
                po_sizes[size] += quantity
            else:
                po_sizes[size] = quantity
    
    return po_sizes

# --- CORRECTION APPLIED HERE (Outside main() for testing/demonstration) ---
pdf_file_path = "Merged_PO_Unknown_BIA_FF_SP_26_INQ_02_PINK KNIT 2_  K43913A6  N51  LBL.CARE_LB 5801 PO 5786464  5786466.msg.pdf"

# Use the correctly named function
email_body_df = extract_email_body_item_data(pdf_file_path)

if email_body_df is not None:
    print("Extracted Email Body Item Data (Color/Garment):")
    print(email_body_df)
   
else:
    print("Failed to extract email body data.")
# --- END CORRECTION ---

def main():
    # Initialize page and session state
    initialize_page()
    initialize_session_state()
    
    # Create sidebar
    create_sidebar()
    
    # Main App Logic
    if st.session_state.checker_name is None:
        st.title("WO & PO Comparison System for LB 5801")
        st.warning("⚠️ Please select a checker name from the sidebar to continue")
        st.stop()

    # Single page layout with all components
    st.title("WO & PO Comparison System for LB 5801")
    
    # Email uploader at the top center
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("### 📧 Email → PDF Merger (Multiple POs Supported)")
        email_file = st.file_uploader("📩 Upload your email file", type=["msg", "eml"], key="email_uploader")
        
        if email_file and st.button("⚡ Convert & Merge", key="convert_merge"):
            result = process_email_to_pdf(email_file)
            
            if result:
                merged_pdf, temp_dir = result
                
                with open(merged_pdf, "rb") as f:
                    st.download_button(
                        label="⬇️ Download Merged PDF",
                        data=f,
                        file_name=os.path.basename(merged_pdf),
                        mime="application/pdf"
                    )
                
                # --- CLEANUP ---
                st.write(f"🧹 Cleaning up temporary directory: `{temp_dir}`")
                shutil.rmtree(temp_dir, ignore_errors=True)
    
    st.markdown("---")
    
    # WO & PO Analysis section
    st.markdown("### 📊 WO & PO Analysis")
    
    # Create two columns for side-by-side uploaders
    col1, col2 = st.columns(2)
    
    # --- FIX IS HERE ---
    # Initialize variables in the outer scope to make them accessible later
    wo_file = None
    merged_po_file = None

    # Left column: WO Upload
    with col1:
        st.markdown("#### Upload WO PDF")
        st.markdown("Drag and drop files here")
        wo_file = st.file_uploader("Browse files", type=['pdf'], key='wo', label_visibility="collapsed")
        st.markdown("Limit 200MB per file • PDF")
        
        # ... (rest of the WO processing code remains the same)
        if wo_file:
            # Process WO
            if st.session_state.wo_data is None or 'last_wo_file' not in st.session_state or st.session_state.last_wo_file != wo_file.name:
                with st.spinner("Processing WO file..."):
                    st.session_state.wo_data = process_wo_file(wo_file)
                    st.session_state.last_wo_file = wo_file.name

            if st.session_state.wo_data:
                st.success("✅ WO processed successfully!")
                
                # Automatically show WO details
                display_wo_details(st.session_state.wo_data, st.session_state.checker_name)
                
                # Display WO Table Data
                st.markdown("---")
                st.subheader("📊 WO Table Data Extraction")
                
                # Extract WO items table and STORE in session state
                with st.spinner("Extracting WO items table..."):
                    wo_items = extract_wo_items_table_enhanced(wo_file)
                    st.session_state.wo_items = wo_items 
                
                if st.session_state.wo_items:
                    st.success(f"✅ Successfully extracted {len(st.session_state.wo_items)} items from WO table")
                    
                    # Display the table
                    st.subheader("WO Items Table")
                    wo_df = pd.DataFrame(st.session_state.wo_items)
                    st.dataframe(wo_df, use_container_width=True, hide_index=True)
                    
                    # Download button for WO items
                    csv = wo_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="⬇️ Download WO Items as CSV",
                        data=csv,
                        file_name="wo_items.csv",
                        mime="text/csv"
                    )
                else:
                    st.warning("⚠️ No WO items found in the table")
                
                # NEW: Display sorted WO size breakdown from Items Table
                st.markdown("---")
                st.subheader("📊 WO Size Breakdown (from Items Table)")
                
                # Extract and sort sizes from WO Items Table
                with st.spinner("Extracting and sorting sizes from WO Items Table..."):
                    sorted_wo_sizes = extract_and_sort_wo_sizes(st.session_state.wo_items)
                
                if sorted_wo_sizes:
                    st.success(f"✅ Successfully extracted and sorted {len(sorted_wo_sizes)} sizes")
                    
                    # Display the sorted table
                    size_df = pd.DataFrame(sorted_wo_sizes)
                    st.dataframe(size_df, use_container_width=True, hide_index=True)
                    
                    # Calculate and display total quantity
                    total_qty = sum(item['Order Quantity'] for item in sorted_wo_sizes)
                    st.metric("Total WO Quantity", total_qty)
                    
                    # Download button for sorted size breakdown
                    csv = size_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="⬇️ Download Sorted WO Size Breakdown as CSV",
                        data=csv,
                        file_name="wo_size_breakdown_sorted.csv",
                        mime="text/csv"
                    )
                else:
                    st.warning("⚠️ No size data available in WO Items Table")
                
                # Extract size breakdown table and store in session state (but don't display it)
                with st.spinner("Extracting size breakdown table..."):
                    size_breakdown = extract_size_breakdown_table_robust(wo_file)
                    st.session_state.wo_size_breakdown = size_breakdown  # Store in session state

    # Right column: PO Upload
    with col2:
        st.markdown("#### Upload Merged PO PDF")
        st.markdown("Drag and drop file here")
        merged_po_file = st.file_uploader("Browse files", type=['pdf'], key='merged_po', label_visibility="collapsed")
        st.markdown("Limit 200MB per file • PDF")
        
        # ... (rest of the PO processing code remains the same)
        # Automatically process PO when file is uploaded
        if merged_po_file:
            if merged_po_file.size == 0:
                st.error("Uploaded file is empty. Please select a valid file.")
            else:
                with st.spinner("Extracting PO details..."):
                    # Add debugging option
                    if st.checkbox("Show debug info", key="po_debug"):
                        display_email_po_debug_info(merged_po_file)
                    
                    # Extract and display the email subject line
                    with pdfplumber.open(merged_po_file) as pdf:
                        first_page_text = pdf.pages[0].extract_text() or ""
                        # Look for subject after "PO #" or "PO "
                        subject_match = re.search(r'PO\s*#?\s*(.*?)(?:\n|Factory\s*Code:|COO:)', first_page_text, re.IGNORECASE | re.DOTALL)
                        if subject_match:
                            subject = subject_match.group(1).strip()
                            st.markdown("### 📧 Email Subject")
                            st.info(subject)
                            
                            # Extract all numbers separated by "/"
                            po_numbers = re.findall(r'(\d+)\s*/\s*', subject)
                            if po_numbers:
                                st.markdown("### 📋 Extracted PO Numbers")
                                # Display all extracted numbers
                                for i, po_num in enumerate(po_numbers, 1):
                                    st.write(f"{i}. {po_num}")
                                
                                # Also show as a comma-separated list
                                po_list_str = ", ".join(po_numbers)
                                st.text(f"All PO Numbers: {po_list_str}")
                    
                    po_list = extract_merged_po_details(merged_po_file)
                    st.session_state.po_data = po_list
                    
                    # Show PO numbers found only in subject line
                    email_po_numbers = extract_po_numbers_from_email_body(merged_po_file)
                    if email_po_numbers:
                        st.info(f"📧 Found {len(email_po_numbers)} PO numbers in subject line:")
                        # Display all PO numbers in a clean format
                        po_text = ", ".join(email_po_numbers)
                        st.text(f"Subject POs: {po_text}")
                    
                    # Show matching PO numbers if found
                    matching_pos = [po for po in po_list if po.get('email_po_number')]
                    if matching_pos:
                        st.success(f"✅ Found {len(matching_pos)} PO(s) matching email PO numbers")
                        
                        # Display each match in the specified format
                        for po in matching_pos:
                            st.write(f"PDF PO {po['po_number']} matches Email PO {po['email_po_number']}")
                    elif po_list:
                        st.warning("⚠️ No PO numbers from PDF match those in the email subject")
                    
                    # Display PO details
                    display_merged_po_results(po_list)
                    
                    # NEW: Display extracted email body table data
                    st.markdown("---")
                    st.subheader("📋 Email Body Item Data (Color/Garment)")
                    
                    # Use the new, correct extraction function
                    email_items_df = extract_email_body_item_data(merged_po_file)
                    
                    if email_items_df is not None:
                        st.dataframe(email_items_df, use_container_width=True)
                        st.success(f"✅ Successfully extracted {len(email_items_df)} item rows.")
                        
                        csv = email_items_df.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="⬇️ Download Email Body Items as CSV",
                            data=csv,
                            file_name="email_body_items.csv",
                            mime="text/csv"
                        )
                    else:
                        st.warning("⚠️ Could not extract item data from email body table.")


                    # NEW: Store PO size breakdowns in session state for each PO
                    for po in po_list:
                        po_number = po.get('po_number', '')
                        if po_number:
                            # Calculate size breakdown for this specific PO
                            po_sizes = {}
                            for item in po.get('items', []):
                                size = item.get('size', '').strip()
                                if not size:
                                    continue
                                    
                                quantity = 0
                                if item.get('quantity') and item['quantity'] != '':
                                    try:
                                        quantity = float(item['quantity'])
                                    except (ValueError, TypeError):
                                        quantity = 0
                                
                                if size in po_sizes:
                                    po_sizes[size] += quantity
                                else:
                                    po_sizes[size] = quantity
                            
                            # Store in session state with a unique key for each PO
                            st.session_state[f"po_size_breakdown_{po_number}"] = po_sizes

    # --- NEW SECTION FOR GARMENT DESCRIPTION TABLE ---
    # This is now outside the column blocks, so it can access merged_po_file
    if merged_po_file and st.session_state.po_data:
        st.markdown("---")
        st.markdown("### 📋 Garment Description Table (PO Attachments)")
        
        # Extract the garment description table
        with st.spinner("Extracting Garment description table..."):
            garment_df = extract_garment_description_table(merged_po_file)
        
        if garment_df is not None and not garment_df.empty:
            st.success("✅ Garment description table extracted successfully!")
            
            # Get the first PO number from the extracted PO data
            first_po_number = None
            if st.session_state.po_data and len(st.session_state.po_data) > 0:
                first_po_number = st.session_state.po_data[0].get('po_number', '')
            
            if first_po_number:
                # Filter the garment description table by PO number
                filtered_garment_df = filter_garment_description_by_po(garment_df, first_po_number)
                
                if not filtered_garment_df.empty:
                    st.info(f"Showing Garment Description for PO: {first_po_number}")
                    st.dataframe(filtered_garment_df, use_container_width=True)
                    
                    # Download button for filtered garment description
                    csv = filtered_garment_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="⬇️ Download Filtered Garment Description as CSV",
                        data=csv,
                        file_name=f"garment_description_{first_po_number}.csv",
                        mime="text/csv"
                    )
                else:
                    st.warning(f"No matching Garment Description found for PO: {first_po_number}")
                    
                    # Show the full table for reference
                    st.info("Showing full Garment Description table for reference:")
                    st.dataframe(garment_df, use_container_width=True)
            else:
                st.warning("No PO number available for filtering. Showing full table:")
                st.dataframe(garment_df, use_container_width=True)
        else:
            st.warning("No Garment Description table found in the PDF")

    # --- DETAILED FIELD COMPARISON ---
    # Only show comparisons if both WO and PO data are available
    if st.session_state.wo_data and st.session_state.po_data:
        st.markdown("---")
        st.subheader("🔍 WO & PO Detailed Field Comparison")
        
        # Create a dictionary to map PO numbers to PO data for easier lookup
        po_dict = {}
        for po in st.session_state.po_data:
            po_number = str(po.get('po_number', ''))
            if po_number:
                po_dict[po_number] = po
        
        # Get WO PO number
        wo_po_number = str(st.session_state.wo_data.get('po_number', ''))
        
        # Find the matching PO number
        matched_po_number = ""
        if wo_po_number in po_dict:
            matched_po_number = wo_po_number
        else:
            # Try partial match
            for po_num in po_dict.keys():
                if wo_po_number in po_num:
                    matched_po_number = po_num
                    break
        
        # Display the comparison table
        display_detailed_comparison_table(st.session_state.wo_data, st.session_state.po_data)
        
        # If we have a matched PO, also display the size comparison
        if matched_po_number and 'wo_items' in st.session_state and st.session_state.wo_items:
            st.markdown("---")
            st.markdown("### 🔍 WO & PO Size Comparison")
            
            # Call the size comparison function
            display_size_comparison_for_matched_po(
                st.session_state.wo_items,  # Pass WO items directly
                st.session_state.po_data,
                matched_po_number
            )
        else:
                st.warning("⚠️ WO Items data not available. Please upload a WO file first.")
    else:
            st.warning("⚠️ Could not perform size comparison because no matching PO was found.")

if __name__ == "__main__":
    main()