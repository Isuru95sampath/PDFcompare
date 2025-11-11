import streamlit as st
import os
import tempfile
import shutil
import sys
import pandas as pd

# Import modules
from ui_components import initialize_page, initialize_session_state, create_sidebar, display_wo_details
from email_processor import process_email_to_pdf
from wo_extractor import process_wo_file, extract_wo_items_table_enhanced, extract_size_breakdown_table_robust
from po_extractor import extract_merged_po_details, display_merged_po_results  # Import PO functions

def main():
    # Initialize page and session state
    initialize_page()
    initialize_session_state()
    
    # Create sidebar and get selected page
    page = create_sidebar()
    
    # Main App Logic
    if st.session_state.checker_name is None:
        st.title("WO & PO Comparison System for LB 5801")
        st.warning("⚠️ Please select a checker name from the sidebar to continue")
        st.stop()

    # PAGE 1: Merge PO - Email to PDF Merger
    if page == "Merge PO":
        st.title("📧 Email → PDF Merger (Multiple POs Supported)")

        email_file = st.file_uploader("📩 Upload your email file", type=["msg", "eml"])

        if email_file and st.button("⚡ Convert & Merge"):
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

    # PAGE 2: Combined WO & PO Analysis (Side by Side Uploaders)
    elif page == "WO & PO Analysis":
        st.title("📊 WO & PO Analysis")
        
        # Create two columns for side-by-side uploaders
        col1, col2 = st.columns(2)
        
        # Left column: WO Upload
        with col1:
            st.markdown("### Upload WO PDF")
            st.markdown("Drag and drop files here")
            wo_file = st.file_uploader("Browse files", type=['pdf'], key='wo', label_visibility="collapsed")
            st.markdown("Limit 200MB per file • PDF")
            
            if wo_file:
                # Process WO
                if st.session_state.wo_data is None or 'last_wo_file' not in st.session_state or st.session_state.last_wo_file != wo_file.name:
                    with st.spinner("Processing WO file..."):
                        st.session_state.wo_data = process_wo_file(wo_file)
                        st.session_state.last_wo_file = wo_file.name

                if st.session_state.wo_data:
                    st.success("✅ WO processed successfully!")
                    
                    # Add a button to show WO details
                    if st.button("Show WO Details", key="show_wo"):
                        display_wo_details(st.session_state.wo_data, st.session_state.checker_name)
                        
                        # Display WO Table Data
                        st.markdown("---")
                        st.subheader("📊 WO Table Data Extraction")
                        
                        # Extract WO items table
                        with st.spinner("Extracting WO items table..."):
                            wo_items = extract_wo_items_table_enhanced(wo_file)
                        
                        if wo_items:
                            st.success(f"✅ Successfully extracted {len(wo_items)} items from WO table")
                            
                            # Display the table
                            st.subheader("WO Items Table")
                            wo_df = pd.DataFrame(wo_items)
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
                        
                        # Extract size breakdown table
                        with st.spinner("Extracting size breakdown table..."):
                            size_breakdown = extract_size_breakdown_table_robust(wo_file)
                        
                        if size_breakdown:
                            st.success(f"✅ Successfully extracted size breakdown for {len(size_breakdown)} sizes")
                            
                            # Display the table
                            size_df = pd.DataFrame(size_breakdown)
                            st.dataframe(size_df, use_container_width=True, hide_index=True)
                            
                            # Calculate and display total quantity
                            total_qty = sum(item['Order Quantity'] for item in size_breakdown)
                            st.metric("Total Quantity", total_qty)
                            
                            # Download button for size breakdown
                            csv = size_df.to_csv(index=False).encode('utf-8')
                            st.download_button(
                                label="⬇️ Download Size Breakdown as CSV",
                                data=csv,
                                file_name="size_breakdown.csv",
                                mime="text/csv"
                            )
                        else:
                            st.warning("⚠️ No size breakdown data found in PDF")
        
        # Right column: Merged PO Upload
        with col2:
            st.markdown("### Upload Merged PO PDF")
            st.markdown("Drag and drop file here")
            merged_po_file = st.file_uploader("Browse files", type=['pdf'], key='merged_po', label_visibility="collapsed")
            st.markdown("Limit 200MB per file • PDF")
            
            if merged_po_file and st.button("Extract PO Details", key="extract_po"):
                with st.spinner("Extracting PO details..."):
                    po_list = extract_merged_po_details(merged_po_file)
                    st.session_state.po_data = po_list
                    display_merged_po_results(po_list)
            elif st.session_state.po_data and not merged_po_file:
                # Display previously extracted data if available
                st.success("✅ PO details already extracted!")
                if st.button("Show PO Details", key="show_po"):
                    display_merged_po_results(st.session_state.po_data)
        
        # Display a comparison section if both WO and PO data are available
        if st.session_state.wo_data and st.session_state.po_data:
            st.markdown("---")
            st.subheader("🔍 WO & PO Comparison")
            
            # Create a comparison table
            comparison_data = []
            
            # Get WO size breakdown
            wo_file = st.session_state.get('last_wo_file')
            if wo_file:
                wo_size_breakdown = extract_size_breakdown_table_robust(wo_file)
                
                # Create a dictionary of WO sizes and quantities
                wo_sizes = {item['Size']: item['Order Quantity'] for item in wo_size_breakdown}
                
                # Get PO sizes and quantities
                po_sizes = {}
                for po in st.session_state.po_data:
                    for item in po['items']:
                        size = item['size']
                        quantity = int(item['quantity'])
                        if size in po_sizes:
                            po_sizes[size] += quantity
                        else:
                            po_sizes[size] = quantity
                
                # Create comparison data
                all_sizes = set(list(wo_sizes.keys()) + list(po_sizes.keys()))
                for size in sorted(all_sizes):
                    wo_qty = wo_sizes.get(size, 0)
                    po_qty = po_sizes.get(size, 0)
                    diff = wo_qty - po_qty
                    comparison_data.append({
                        'Size': size,
                        'WO Quantity': wo_qty,
                        'PO Quantity': po_qty,
                        'Difference': diff
                    })
                
                # Display comparison table
                if comparison_data:
                    comparison_df = pd.DataFrame(comparison_data)
                    st.dataframe(comparison_df, use_container_width=True)
                    
                    # Download button for comparison data
                    csv = comparison_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="⬇️ Download Comparison as CSV",
                        data=csv,
                        file_name="wo_po_comparison.csv",
                        mime="text/csv"
                    )
                else:
                    st.warning("⚠️ No data available for comparison")

if __name__ == "__main__":
    main()