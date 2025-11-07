import streamlit as st  # Added this import
import pandas as pd
import re
from fuzzywuzzy import fuzz

def enhanced_quantity_matching(wo_items, po_details, tolerance=0, excel_style=None):
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
            
            # If Style 2 is empty and we have an Excel style, use it for comparison
            if not pstyle and excel_style:
                pstyle = excel_style
            
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
                # If Style 2 is empty and we have an Excel style, use it for comparison
                if not pstyle and excel_style:
                    if wstyle == excel_style:
                        score += 1
                elif wstyle == pstyle:
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
            
            # If Style 2 is empty and we have an Excel style, use it for comparison
            if not pstyle and excel_style:
                pstyle = excel_style
            
            qty_match = "👍 Yes" if abs(pq - wq) <= tolerance else "❌ No"
            size_match = "👍 Yes" if ws == ps else "❌ No"
            colour_match = "👍 Yes" if wc == pc else "❌ No"
            # If Style 2 is empty and we have an Excel style, use it for comparison
            if not po.get("Style 2", "").strip() and excel_style:
                style_match = "👍 Yes" if wstyle == excel_style else "❌ No"
            else:
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
            
            # If Style 2 is empty and we have an Excel style, use it for comparison
            if not pstyle and excel_style:
                pstyle = excel_style
            
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

def fill_empty_style_2_from_excel(matched_items, mismatched_items, processed_excel_data):
    """
    Fill empty Style 2 values in matched and mismatched items with style number from Excel
    
    Args:
        matched_items: List of matched items from comparison
        mismatched_items: List of mismatched items from comparison
        processed_excel_data: DataFrame containing processed Excel data
    
    Returns:
        Updated matched_items and mismatched_items with empty Style 2 filled from Excel
    """
    # If no processed Excel data, return original items
    if processed_excel_data is None or processed_excel_data.empty:
        return matched_items, mismatched_items
    
    # Extract the first style number from Excel data
    excel_style = None
    for _, row in processed_excel_data.iterrows():
        style = str(row.get('Style', '')).strip()
        if style:
            excel_style = style
            break
    
    # If no style found in Excel, return original items
    if not excel_style:
        return matched_items, mismatched_items
    
    # Update matched items with Excel style only if Style 2 is empty
    updated_matched = []
    for item in matched_items:
        updated_item = item.copy()
        if not updated_item.get('Style 2', ''):
            updated_item['Style 2'] = excel_style
        updated_matched.append(updated_item)
    
    # Update mismatched items with Excel style only if Style 2 is empty
    updated_mismatched = []
    for item in mismatched_items:
        updated_item = item.copy()
        if not updated_item.get('Style 2', ''):
            updated_item['Style 2'] = excel_style
        updated_mismatched.append(updated_item)
    
    return updated_matched, updated_mismatched  

def compare_codes(po_details, wo_items, po_product_codes_from_item=None):
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
    
    # Add PO product codes from Item column if available
    if po_product_codes_from_item:
        for code in po_product_codes_from_item:
            if code:
                cleaned_code = str(code).strip().upper()
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

def update_po_details_with_excel_styles(po_details, processed_data):
    """Update PO details with style numbers from processed Excel data - only at the very top"""
    if processed_data.empty:
        return po_details
    
    # Get the first style number from processed Excel data
    first_style = None
    for _, row in processed_data.iterrows():
        style = str(row.get('Style', '')).strip()
        if style:
            first_style = style
            break  # Only get the first style number
    
    # If no style found, return original PO details
    if not first_style:
        return po_details
    
    # DO NOT update any PO items with style number
    # Just return the original PO details unchanged
    # The style number will be handled separately in main.py
    return po_details

def get_excel_style_number(processed_data):
    """Get the first style number from processed Excel data"""
    if processed_data.empty:
        return None
    
    # Get the first style number from processed Excel data
    for _, row in processed_data.iterrows():
        style = str(row.get('Style', '')).strip()
        if style:
            return style
    
    return None

def update_matched_items_with_excel_styles(matched_items, mismatched_items, processed_data):
    """Update matched and mismatched items with style numbers from processed Excel data - DISABLED"""
    # Return items unchanged - style number is displayed separately at the top
    return matched_items, mismatched_items

def extract_style_from_excel(processed_data):
    """Extract the first style number from processed Excel data"""
    if processed_data.empty:
        return None
    
    # Get the first style number from processed Excel data
    for _, row in processed_data.iterrows():
        style = str(row.get('Style', '')).strip()
        if style:
            return style
    
    return None

def combine_wo_and_excel_data(wo_df, excel_df):
    """Combine WO and Excel data into a single table with paired columns and comparison results"""
    try:
        if wo_df.empty and excel_df.empty:
            return pd.DataFrame()
        
        # Create a copy of the dataframes to avoid modifying the originals
        wo_combined = wo_df.copy()
        excel_combined = excel_df.copy()
        
        # Remove WO Product Code column if it exists
        if 'WO Product Code' in wo_combined.columns:
            wo_combined = wo_combined.drop(columns=['WO Product Code'])
        
        # Create a combined dataframe
        combined_df = pd.DataFrame()
        
        # Add WO columns with prefix (excluding WO Product Code)
        for col in wo_combined.columns:
            combined_df[f"WO {col}"] = None
        
        # Add Excel columns with prefix
        for col in excel_combined.columns:
            combined_df[f"Excel {col}"] = None
        
        # Standardize column names for matching
        wo_col_mapping = {
            'Style': 'Style',
            'WO Colour Code': 'Colour Code',
            'Size 1': 'Size',
            'Article': 'Article'
        }
        
        excel_col_mapping = {
            'Style': 'Style',
            'Colour Code': 'Colour Code',
            'Size': 'Size',
            'Article': 'Article'
        }
        
        # Create standardized dataframes for matching
        wo_standard = pd.DataFrame()
        for wo_col, std_col in wo_col_mapping.items():
            if wo_col in wo_combined.columns:
                wo_standard[std_col] = wo_combined[wo_col]
        
        excel_standard = pd.DataFrame()
        for excel_col, std_col in excel_col_mapping.items():
            if excel_col in excel_combined.columns:
                excel_standard[std_col] = excel_combined[excel_col]
        
        # Create a key for matching
        wo_standard['key'] = wo_standard['Style'].astype(str) + '_' + wo_standard['Colour Code'].astype(str) + '_' + wo_standard['Size'].astype(str)
        excel_standard['key'] = excel_standard['Style'].astype(str) + '_' + excel_standard['Colour Code'].astype(str) + '_' + excel_standard['Size'].astype(str)
        
        # Create a dictionary to map keys to row indices
        wo_key_to_idx = {key: idx for idx, key in wo_standard['key'].items()}
        excel_key_to_idx = {key: idx for idx, key in excel_standard['key'].items()}
        
        # Find matching keys
        matching_keys = set(wo_key_to_idx.keys()) & set(excel_key_to_idx.keys())
        
        # Process matching rows
        for key in matching_keys:
            wo_idx = wo_key_to_idx[key]
            excel_idx = excel_key_to_idx[key]
            
            # Add WO data
            for col in wo_combined.columns:
                if col != 'WO Product Code':
                    combined_df.at[wo_idx, f"WO {col}"] = wo_combined.at[wo_idx, col]
            
            # Add Excel data
            for col in excel_combined.columns:
                combined_df.at[wo_idx, f"Excel {col}"] = excel_combined.at[excel_idx, col]
        
        # Process WO-only rows
        wo_only_keys = set(wo_key_to_idx.keys()) - matching_keys
        for key in wo_only_keys:
            wo_idx = wo_key_to_idx[key]
            for col in wo_combined.columns:
                if col != 'WO Product Code':
                    combined_df.at[wo_idx, f"WO {col}"] = wo_combined.at[wo_idx, col]
        
        # REMOVED: Process Excel-only rows section
        # This section has been removed to exclude Excel-only rows from the combined dataframe
        
        # Clean up the dataframe
        combined_df = combined_df.dropna(how='all').reset_index(drop=True)
        
        # Clean retail columns
        retail_columns = ['WO Retail US', 'WO Retail CA', 'Excel Retail US', 'Excel Retail CA']
        for col in retail_columns:
            if col in combined_df.columns:
                from excel_utils import clean_retail_value
                combined_df[col] = combined_df[col].apply(clean_retail_value)
        
        # Clean decimal values from SKU, Article, and Quantity columns
        columns_to_clean = [
            'WO SKU', 'Excel SKU',
            'WO Article', 'Excel Article',
        ]
        
        for col in columns_to_clean:
            if col in combined_df.columns:
                from excel_utils import clean_decimal_values
                combined_df[col] = combined_df[col].apply(clean_decimal_values)
        
        # Create comparison columns
        combined_df["Overall Match"] = None
        
        # Compare Color Codes (WO Colour Code vs Excel Colour Code)
        combined_df["Match Color"] = None
        for idx, row in combined_df.iterrows():
            # Fixed syntax error here
            if "WO WO Colour Code" in combined_df.columns:
                wo_color = row.get("WO WO Colour Code", None)
            else:
                wo_color = row.get("WO Colour Code", None)
            
            excel_color = row.get("Excel Colour Code", None)
            
            if pd.isna(wo_color) and pd.isna(excel_color):
                combined_df.at[idx, "Match Color"] = "👍 Yes"
            elif pd.isna(wo_color) or pd.isna(excel_color):
                combined_df.at[idx, "Match Color"] = "❌"
            else:
                wo_str = str(wo_color).strip().upper()
                excel_str = str(excel_color).strip().upper()
                combined_df.at[idx, "Match Color"] = "👍 Yes" if wo_str == excel_str else "❌"
        
        # Compare Sizes (WO Size 1 vs Excel Size)
        combined_df["Match Size"] = None
        for idx, row in combined_df.iterrows():
            wo_size = row.get("WO Size 1", None)
            excel_size = row.get("Excel Size", None)
            
            if pd.isna(wo_size) and pd.isna(excel_size):
                combined_df.at[idx, "Match Size"] = "👍 Yes"
            elif pd.isna(wo_size) or pd.isna(excel_size):
                combined_df.at[idx, "Match Size"] = "❌"
            else:
                wo_str = str(wo_size).strip().upper()
                excel_str = str(excel_size).strip().upper()
                combined_df.at[idx, "Match Size"] = "👍 Yes" if wo_str == excel_str else "❌"
        
        # Compare Articles (WO Article vs Excel Article)
        combined_df["Match Article"] = None
        for idx, row in combined_df.iterrows():
            wo_article = row.get("WO Article", None)
            excel_article = row.get("Excel Article", None)
            
            if pd.isna(wo_article) and pd.isna(excel_article):
                combined_df.at[idx, "Match Article"] = "👍 Yes"
            elif pd.isna(wo_article) or pd.isna(excel_article):
                combined_df.at[idx, "Match Article"] = "❌"
            else:
                # Clean and compare article numbers
                from excel_utils import clean_decimal_values
                wo_str = clean_decimal_values(wo_article).strip().upper()
                excel_str = clean_decimal_values(excel_article).strip().upper()
                combined_df.at[idx, "Match Article"] = "👍 Yes" if wo_str == excel_str else "❌"
        
        # Compare SKU (with leading zero removal)
        combined_df["Match SKU"] = None
        for idx, row in combined_df.iterrows():
            wo_sku = row.get("WO SKU", None)
            excel_sku = row.get("Excel SKU", None)
            
            if pd.isna(wo_sku) and pd.isna(excel_sku):
                combined_df.at[idx, "Match SKU"] = "👍 Yes"
            elif pd.isna(wo_sku) or pd.isna(excel_sku):
                combined_df.at[idx, "Match SKU"] = "❌"
            else:
                from excel_utils import clean_decimal_values, remove_leading_zeros
                wo_str = remove_leading_zeros(clean_decimal_values(wo_sku)).strip().upper()
                excel_str = remove_leading_zeros(clean_decimal_values(excel_sku)).strip().upper()
                combined_df.at[idx, "Match SKU"] = "👍 Yes" if wo_str == excel_str else "❌"
        
        
        # Calculate Overall Match - UPDATED LOGIC
        for idx, row in combined_df.iterrows():
            # Get all match column values
            match_color = row.get("Match Color", None)
            match_size = row.get("Match Size", None)
            match_article = row.get("Match Article", None)
            match_sku = row.get("Match SKU", None)
           
            
            # Check if we have data in both WO and Excel
            has_wo_data = any(not pd.isna(row.get(f"WO {col}", None)) for col in ['Colour Code', 'Size 1', 'Quantity', 'SKU', 'Article'] if f"WO {col}" in row.index or "WO WO Colour Code" in row.index)
            has_excel_data = any(not pd.isna(row.get(f"Excel {col}", None)) for col in ['Colour Code', 'Size', 'Quantity', 'SKU', 'Article'] if f"Excel {col}" in row.index)
            
            if has_wo_data and has_excel_data:
                # Collect all match results (only those that are not None)
                all_matches = []
                if match_color is not None:
                    all_matches.append(match_color)
                if match_size is not None:
                    all_matches.append(match_size)
                if match_article is not None:
                    all_matches.append(match_article)
                if match_sku is not None:
                    all_matches.append(match_sku)
              
                
                # If ALL comparisons are ✅, then Full Match, otherwise Mismatch
                if all_matches and all(val == "👍 Yes" for val in all_matches):
                    combined_df.at[idx, "Overall Match"] = "✅ Full Match"
                else:
                    combined_df.at[idx, "Overall Match"] = "❌ Mismatch"
            elif has_wo_data and not has_excel_data:
                combined_df.at[idx, "Overall Match"] = "⚠️ Missing Excel Data"
            elif not has_wo_data and has_excel_data:
                combined_df.at[idx, "Overall Match"] = "⚠️ Missing WO Data"
            else:
                combined_df.at[idx, "Overall Match"] = "⚠️ No Data"
        
        # Sort by size order
        size_order = {
            "XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5, "XXXL": 6,
            "P": 7, "G": 8, "XG": 9, "XXG": 10
        }
        
        def get_size_order(size_str):
            if pd.isna(size_str):
                return 999
            size_str = str(size_str).strip().upper()
            if size_str in size_order:
                return size_order[size_str]
            if "/" in size_str:
                first_part = size_str.split("/")[0].strip()
                if first_part in size_order:
                    return size_order[first_part]
            for size in size_order.keys():
                if size in size_str:
                    return size_order[size]
            return 999
        
        combined_df['size_order'] = combined_df.apply(
            lambda row: get_size_order(row.get('WO Size 1', row.get('Excel Size', None))), 
            axis=1
        )
        combined_df = combined_df.sort_values('size_order').reset_index(drop=True)
        combined_df = combined_df.drop(columns=['size_order'])
        
        # Fix column names
        if "WO WO Colour Code" in combined_df.columns:
            combined_df = combined_df.rename(columns={"WO WO Colour Code": "WO Colour Code"})
        
        # Reorder columns for better readability
        desired_column_order = []
        
        # WO columns
        wo_columns_to_include = ["WO Colour Code", "WO Size 1", "WO Quantity", "WO SKU", "WO Article"]
        for col in wo_columns_to_include:
            if col in combined_df.columns:
                desired_column_order.append(col)
        
        for col in combined_df.columns:
            if col.startswith("WO ") and col not in wo_columns_to_include:
                desired_column_order.append(col)
        
        # Excel columns
        excel_columns_to_include = ["Excel Colour Code", "Excel Size", "Excel Quantity", "Excel SKU", "Excel Article"]
        for col in excel_columns_to_include:
            if col in combined_df.columns:
                desired_column_order.append(col)
        
        for col in combined_df.columns:
            if col.startswith("Excel ") and col not in excel_columns_to_include:
                desired_column_order.append(col)
        
        # Match columns
        match_columns = ["Match Color", "Match Size", "Match Quantity", "Match SKU", "Match Article"]
        for col in match_columns:
            if col in combined_df.columns:
                desired_column_order.append(col)
        
        for col in combined_df.columns:
            if col.startswith("Match ") and col not in match_columns:
                desired_column_order.append(col)
        
        # Overall Match at the end
        if "Overall Match" in combined_df.columns:
            if "Overall Match" in desired_column_order:
                desired_column_order.remove("Overall Match")
            desired_column_order.append("Overall Match")
        
        # Add remaining columns
        for col in combined_df.columns:
            if col not in desired_column_order:
                desired_column_order.append(col)
        
        combined_df = combined_df[desired_column_order]
        
        return combined_df
    except Exception as e:
        st.error(f"Error combining WO and Excel data: {e}")
        return pd.DataFrame()

def update_so_color_display(so_numbers, wo_items):
    """
    Display WO Color Codes vs SO Numbers Comparison table with each WO item in a separate row.
    
    Args:
        so_numbers: List of SO numbers
        wo_items: List of WO items with color codes
        
    Returns:
        DataFrame with comparison results
    """
    # Extract WO color codes (keep duplicates as they appear in the WO)
    wo_color_codes = []
    for item in wo_items:
        color_code = item.get('WO Colour Code', '')
        if color_code:
            wo_color_codes.append(color_code)
    
    # Create comparison data with each WO item in a separate row
    comparison_data = []
    
    # Process each WO item
    for i, item in enumerate(wo_items):
        color_code = item.get('WO Colour Code', '')
        if not color_code:
            continue  # Skip items without color code
        
        # Find matching SO numbers for this color
        matching_so_numbers = []
        for so_number in so_numbers:
            if so_number and color_code.lower() in so_number.lower():
                matching_so_numbers.append(so_number)
        
        # Determine match status
        if matching_so_numbers:
            status = "✅ Match"
        else:
            status = "⚠️ No SO Number"
        
        # Create a row for this WO item
        comparison_data.append({
            "WO Color Code": color_code,
            "SO Numbers": ", ".join(matching_so_numbers) if matching_so_numbers else "",
            "Status": status,
            "WO Item": f"Item {i+1}"  # Add item number for reference
        })
    
    # Create DataFrame
    df = pd.DataFrame(comparison_data)
    
    # Display summary counts
    total_items = len(df)
    matched_items = len(df[df["Status"] == "✅ Match"])
    no_so_items = len(df[df["Status"] == "⚠️ No SO Number"])
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total WO Items", total_items)
    with col2:
        st.metric("Matched", matched_items)
    with col3:
        st.metric("No SO", no_so_items)
    
    # Display the table
    st.dataframe(df, use_container_width=True)
    
    # Display summary
    st.markdown(f"""
    <div class="alert-info">
        <strong>Summary:</strong> {total_items} WO items — ✅ {matched_items} Matched, ⚠️ {no_so_items} No SO Number
    </div>
    """, unsafe_allow_html=True)
    
    return df

def clean_product_code(code):
    """Remove '-VSBA' suffix from product codes and clean up"""
    if not code:
        return ""
    # Convert to string and uppercase
    code_str = str(code).strip().upper()
    # Remove -VSBA suffix
    if code_str.endswith("-VSBA"):
        code_str = code_str[:-5].strip()  # Remove the last 5 characters and strip
    # Also handle case without dash
    if code_str.endswith("VSBA"):
        code_str = code_str[:-4].strip()  # Remove the last 4 characters and strip
    return code_str

def sort_items_by_size(items):
    size_order = {"XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5}
    def get_size_key(item):
        size = item.get("WO Size", item.get("PO Size", "")).strip().upper()
        return size_order.get(size, 99)
    return sorted(items, key=get_size_key)