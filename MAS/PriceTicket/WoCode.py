import pdfplumber
import pandas as pd
import re
import streamlit as st

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


# --- Streamlit UI ---
st.title("MAS WO PDF Extractor 📄")

uploaded_files = st.file_uploader("Upload one or more WO PDFs", type=["pdf"], accept_multiple_files=True)
if uploaded_files:
    all_dfs = []
    individual_summaries = []
    
    with st.spinner("Processing PDF(s) and extracting data..."):
        for file in uploaded_files:
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
        # Display individual tables for each WO
        st.write("### Individual WO Extractions")
        
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
                    key=f"download_{i}",
                    use_container_width=True
                )

        # Combined data section
        st.write("### Combined Summary")
        combined_df = pd.concat(all_dfs, ignore_index=True)
        
        # Summary table of all files
        summary_df = pd.DataFrame(individual_summaries)
        st.dataframe(summary_df, use_container_width=True)
        
        # Overall totals
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Files", len(all_dfs))
        with col2:
            st.metric("Total Items", len(combined_df))
        with col3:
            total_qty = pd.to_numeric(combined_df["Quantity"], errors="coerce").fillna(0).astype(int).sum()
            st.metric("Combined Quantity", f"{total_qty:,}")
        with col4:
            st.metric("All Unique Styles", combined_df["Style"].nunique())

        # Combined download option
        csv_combined = combined_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Download All Combined CSV",
            data=csv_combined,
            file_name="all_wo_combined_data.csv",
            mime="text/csv",
            use_container_width=True
        )
        
        # Optional: Show combined data table in expander
        with st.expander("View Combined Data Table", expanded=False):
            st.dataframe(combined_df, use_container_width=True)
    else:
        st.warning("No data extracted from uploaded PDFs.")
