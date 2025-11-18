import streamlit as st
import pandas as pd
import re
from typing import List, Dict, Any, Tuple, Optional

def normalize_po_number(po_number: str) -> str:
    """
    Normalizes a PO number by removing common prefixes and cleaning it.
    Examples: 
    - 'BFF5791097' -> '5791097'
    - '5791097' -> '5791097'
    - 'PO5791097' -> '5791097'
    """
    if not po_number:
        return ""
    
    po_clean = str(po_number).strip().upper()
    
    # Remove common prefixes
    prefixes = ['BFF', 'PO', 'P.O.', 'PO#', 'P.O#']
    for prefix in prefixes:
        if po_clean.startswith(prefix):
            po_clean = po_clean[len(prefix):].strip()
    
    # Remove any non-digit characters except hyphens
    po_clean = re.sub(r'[^\d-]', '', po_clean)
    
    return po_clean

def find_matching_po_in_list(wo_po_number: str, po_list: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """
    Finds the matching PO from the PO list based on WO PO number.
    Returns the complete PO dictionary with all items and details.
    
    Matching priority:
    1. Exact match of normalized PO numbers
    2. Partial match (WO PO contained in PDF PO)
    3. Check email_po_number field if available
    """
    if not wo_po_number or not po_list:
        return None
    
    wo_po_normalized = normalize_po_number(wo_po_number)
    
    # Priority 1: Exact match on normalized PO number
    for po in po_list:
        pdf_po_normalized = normalize_po_number(po.get('po_number', ''))
        if wo_po_normalized == pdf_po_normalized:
            return po
    
    # Priority 2: Check email_po_number field
    for po in po_list:
        email_po = po.get('email_po_number', '')
        if email_po:
            email_po_normalized = normalize_po_number(email_po)
            if wo_po_normalized == email_po_normalized:
                return po
    
    # Priority 3: Partial match (WO PO contained in PDF PO)
    for po in po_list:
        pdf_po_normalized = normalize_po_number(po.get('po_number', ''))
        if wo_po_normalized in pdf_po_normalized or pdf_po_normalized in wo_po_normalized:
            return po
    
    return None

def extract_color_code_from_description(description: str) -> str:
    """
    Extracts color code from product description.
    Looks for patterns like: C307, C/307, C1, C/1
    """
    if not description:
        return ""
    
    # Pattern 1: C followed by digits (e.g., C307)
    match1 = re.search(r'\bC(\d+)\b', description)
    if match1:
        return f"C{match1.group(1)}"
    
    # Pattern 2: C/ followed by digits (e.g., C/307)
    match2 = re.search(r'\bC/(\d+)\b', description)
    if match2:
        return f"C/{match2.group(1)}"
    
    return ""

def extract_complete_po_data(po: Dict[str, Any]) -> Dict[str, Any]:
    """
    Extracts ALL relevant fields from a matched PO for comparison.
    Intelligently searches through all items to find the first non-empty value for each field.
    Includes logic to filter out color codes that are part of VSD/VSS fields.
    """
    if not po or not isinstance(po, dict):
        return {}
    
    # Initialize with PO-level data
    po_data = {
        'po_number': str(po.get('po_number', '')),
        'supplier': str(po.get('supplier', '')),
        'delivery_location': str(po.get('delivery_location', '')),
        'customer': str(po.get('customer', '')),
        'address': str(po.get('address', '')),
        'delivery_date': str(po.get('delivery_date', '')),
        'season': str(po.get('season', '')),
        'total_quantity': str(po.get('total_quantity', '')),
    }
    
    # Fields to extract from items (will search all items for first non-empty value)
    item_fields = {
        'factory_id': ['factory_id'],
        'date_of_mfr': ['date_of_mfr'],
        'vss_vsd': ['vsd', 'vss_vsd'],
        'silhouette': ['silhouette'],
        'product_code': ['product_code'],
        'care_instruction': ['care_instruction'],
        'size_id': ['size', 'size_id'],
        'garment_components': ['description', 'garment_components'],
    }
    
    # Get items list
    items = po.get('items', [])
    if not items or not isinstance(items, list):
        for field_name in item_fields.keys():
            po_data[field_name] = ''
        po_data['color_code'] = ''
        return po_data

    # Define a check function inside the main extraction function to avoid VSD/VSS conflict
    def is_valid_po_color_code(code: str, item: Dict[str, Any]) -> bool:
        """Checks if the extracted color code is valid and not part of the VSD/VSS field."""
        if not code:
            return False
        
        # Get the VSD/VSS field from the current item
        vsd_value = str(item.get('vsd', '')).strip() or str(item.get('vss_vsd', '')).strip()
        
        if not vsd_value:
            return bool(code)
            
        # The core logic: "do not get color code as a part of Vss Vsd"
        # We reject the color code if it is a substring of the VSD/VSS value.
        if code in vsd_value:
            return False
            
        return True
    
    # Special handling for color_code - try multiple sources
    color_code_found = ''
    
    for item in items:
        if not isinstance(item, dict):
            continue
        
        # Priority 1: Check 'color_code' field directly
        color_code = str(item.get('color_code', '')).strip()
        if is_valid_po_color_code(color_code, item):
            color_code_found = color_code
            break
        
        # Priority 2: Check 'color_code_from_product' field
        color_code = str(item.get('color_code_from_product', '')).strip()
        if is_valid_po_color_code(color_code, item):
            color_code_found = color_code
            break
        
        # Priority 3: Check 'third_line' field for color code
        third_line = str(item.get('third_line', '')).strip()
        if third_line:
            # Extract color code from third line using regex
            color_match = re.search(r'\b(C\d+|C/\d+)\b', third_line)
            if color_match:
                color_code = color_match.group(1)
                if is_valid_po_color_code(color_code, item):
                    color_code_found = color_code
                    break
        
        # Priority 4: Extract from description
        description = str(item.get('description', '')).strip()
        if description:
            color_code = extract_color_code_from_description(description) 
            if is_valid_po_color_code(color_code, item):
                color_code_found = color_code
                break
    
    # Add the found color code to the PO data
    po_data['color_code'] = color_code_found

    # Get the first non-empty value for other item fields
    for field_name, keys in item_fields.items():
        value_found = ''
        for item in items:
            if not isinstance(item, dict):
                continue
            for key in keys:
                value = str(item.get(key, '')).strip()
                if value:
                    value_found = value
                    break
            if value_found:
                break
        po_data[field_name] = value_found
            
    return po_data

def extract_wo_comparison_data(wo_data: Dict[str, Any]) -> Dict[str, str]:
    """ Extract relevant data from WO for comparison. Ensures all values are strings. """
    if not isinstance(wo_data, dict):
        raise ValueError("WO data is not in the expected format (dictionary).")
    return {
        'po_number': str(wo_data.get('po_number', '')),
        'color_code': str(wo_data.get('color_code', '')),
        'factory_id': str(wo_data.get('factory_id', '')),
        'date_of_mfr': str(wo_data.get('date_of_mfr', '')),
        'vss_vsd': str(wo_data.get('vss_vsd', '')),
        'silhouette': str(wo_data.get('silhouette', '')),
        'product_code': str(wo_data.get('product_code', '')),
        'care_instruction': str(wo_data.get('care_instruction', '')),
        'season': str(wo_data.get('season', '')),
        'quantity': str(wo_data.get('quantity', '')),
        'delivery_date': str(wo_data.get('delivery_date', '')),
        'size_id': str(wo_data.get('size_id', '')),
        'delivery_location': str(wo_data.get('delivery_location', '')),
        'customer': str(wo_data.get('customer', '')),
        'address': str(wo_data.get('address', '')),
        'garment_components': str(wo_data.get('garment_components', '')),
    }

def is_product_code_match(wo_code: str, po_code: str) -> bool:
    """ Checks if a WO product code matches a PO product code (partial match allowed). """
    wo_clean = re.sub(r'[^A-Z0-9]', '', wo_code.upper())
    po_clean = re.sub(r'[^A-Z0-9]', '', po_code.upper())
    
    if wo_clean and po_clean:
        return wo_clean in po_clean or po_clean in wo_clean
    return False

def normalize_color_code(color_code: str) -> str:
    """Normalizes color codes, treating 'C/1' and 'C1' as equivalent by removing the slash."""
    if not color_code:
        return ""
    # Remove '/' if it appears after the initial 'C'
    normalized = color_code.strip().upper().replace('C/', 'C')
    return normalized

def compare_wo_po_data(wo_data: Dict[str, Any], po_list: List[Dict[str, Any]]) -> Tuple[pd.DataFrame, str, str, Dict[str, Any]]:
    """
    Performs the core comparison logic between WO and PO data.
    Returns: comparison_df, wo_po_number, wo_po_number, matched_po_full_details
    """
    if not wo_data:
        raise ValueError("WO data is not available for comparison.")
    if not po_list:
        raise ValueError("PO data is not available for comparison.")

    # Find the matching PO
    wo_po_number = wo_data.get('po_number', '')
    matched_po = find_matching_po_in_list(wo_po_number, po_list)
    
    if not matched_po:
        raise ValueError(f"Could not find a matching PO for WO PO number: {wo_po_number}")

    # Extract comparison data
    wo_comparison_data = extract_wo_comparison_data(wo_data)
    po_comparison_data = extract_complete_po_data(matched_po)
    
    # Fields to compare
    fields_to_compare = [
        'po_number', 'customer', 'delivery_location', 'address', 'delivery_date',
        'season', 'color_code', 'factory_id', 'date_of_mfr', 'vss_vsd', 
        'silhouette', 'product_code', 'care_instruction', 'size_id', 'garment_components',
        'quantity' # Comparing total quantity
    ]
    
    field_comparison = {'Field': [], 'WO Value': [], 'PO Value': [], 'Status': []}
    
    # Create the comparison DataFrame
    for field in fields_to_compare:
        wo_value = wo_comparison_data.get(field, '')
        po_value = po_comparison_data.get(field, '')

        # Determine status
        status = ""
        if field == 'product_code':
            # Special handling for product code
            if is_product_code_match(wo_value, po_value):
                status = "✅ Matched"
            elif not wo_value and not po_value:
                status = "⚪ Both Empty"
            elif wo_value and not po_value:
                status = "❌ Missing in PO"
            elif not wo_value and po_value:
                status = "❌ Missing in WO"
            else:
                status = "❌ Mismatch"
        elif field == 'color_code':
            # Apply custom color code normalization for comparison (C1 == C/1)
            wo_normalized = normalize_color_code(wo_value)
            po_normalized = normalize_color_code(po_value)
            
            if wo_normalized and po_normalized:
                if wo_normalized == po_normalized:
                    status = "✅ Matched"
                else:
                    status = "❌ Mismatch"
            elif not wo_value and not po_value:
                status = "⚪ Both Empty"
            elif wo_value and not po_value:
                status = "❌ Missing in PO"
            elif not wo_value and po_value:
                status = "❌ Missing in WO"
            else:
                status = "❌ Mismatch"
        else:
            # Standard matching for all other fields
            if wo_value == po_value:
                status = "✅ Matched"
            elif not wo_value and not po_value:
                status = "⚪ Both Empty"
            elif wo_value and not po_value:
                status = "❌ Missing in PO"
            elif not wo_value and po_value:
                status = "❌ Missing in WO"
            else:
                status = "❌ Mismatch"

        # Format field name for display
        field_display = field.replace('_', ' ').title()
        field_comparison['Field'].append(field_display)
        field_comparison['WO Value'].append(wo_value if wo_value else '(empty)')
        field_comparison['PO Value'].append(po_value if po_value else '(empty)')
        field_comparison['Status'].append(status)

    comparison_df = pd.DataFrame(field_comparison)
    
    # Return 4 values to match the user's expected assignment
    return comparison_df, wo_po_number, wo_po_number, matched_po


def display_comparison_table(wo_data: Dict[str, Any], po_list: List[Dict[str, Any]]):
    """ Display comprehensive comparison table between WO and PO data. """
    try:
        # Call the newly defined function to get comparison data
        comparison_df, wo_po_number_for_display, matched_po_number, matched_po_full = compare_wo_po_data(wo_data, po_list)
    except ValueError as e:
        st.error(f"❌ Comparison Error: {e}")
        return

    # Display the table
    st.dataframe(comparison_df, use_container_width=True, hide_index=True)
    
    # Download button
    csv = comparison_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="⬇️ Download Comparison Table as CSV",
        data=csv,
        file_name=f"wo_po_field_comparison_{matched_po_number}.csv",
        mime="text/csv",
        key="download_field_comparison"
    )
    # Size comparison (if available)
    if matched_po_full and isinstance(matched_po_full, dict) and 'wo_items' in st.session_state and st.session_state.wo_items:
        st.markdown("---")
        st.markdown(f"### 📊 Size and Quantity Breakdown Comparison")
        display_size_comparison(matched_po_full, st.session_state.wo_items)


def display_size_comparison(matched_po: Dict[str, Any], wo_items: List[Dict], po_number: str):
    """
    Display size and quantity comparison between WO and matched PO.
    Now accepts po_number to create unique keys.
    """
    # Validate matched_po is a dictionary
    if not matched_po or not isinstance(matched_po, dict):
        st.error("❌ Invalid PO data for size comparison")
        return
    
    # The po_number is now passed in, but we can also get it as a fallback
    matched_po_number = po_number or matched_po.get('po_number', '')
    
    if not matched_po_number:
        st.error("❌ Could not extract PO number from matched PO")
        return
    
    # ... (rest of the function remains the same until the download button)
    
    # Extract WO sizes
    wo_sizes = {}
    if wo_items and isinstance(wo_items, list):
        for item in wo_items:
            if not isinstance(item, dict):
                continue
                
            size = (item.get('Size 1') or item.get('Size') or item.get('size') or '').strip()
            if not size:
                continue
                
            quantity = 0
            qty_str = (item.get('Quantity') or item.get('quantity') or '')
            if qty_str:
                try:
                    quantity = float(str(qty_str).replace(',', ''))
                except (ValueError, TypeError):
                    quantity = 0
            
            wo_sizes[size] = wo_sizes.get(size, 0) + quantity
    
    # Extract PO sizes from matched PO
    po_sizes = {}
    po_items = matched_po.get('items', [])
    if po_items and isinstance(po_items, list):
        for item in po_items:
            if not isinstance(item, dict):
                continue
                
            size = str(item.get('size', '')).strip()
            if not size:
                continue
                
            quantity = 0
            if item.get('quantity'):
                try:
                    quantity = float(str(item['quantity']).replace(',', ''))
                except (ValueError, TypeError):
                    quantity = 0
            
            po_sizes[size] = po_sizes.get(size, 0) + quantity
    
    if not wo_sizes and not po_sizes:
        st.warning("⚠️ No size data available for comparison")
        return
    
    # Create comparison data
    size_order = {'XS': 0, 'S': 1, 'M': 2, 'L': 3, 'XL': 4, 'XXL': 5, 'XXXL': 6}
    all_sizes = sorted(
        set(list(wo_sizes.keys()) + list(po_sizes.keys())),
        key=lambda x: size_order.get(x.upper(), 999)
    )
    
    comparison_data = []
    for size in all_sizes:
        wo_qty = wo_sizes.get(size, 0)
        po_qty = po_sizes.get(size, 0)
        difference = wo_qty - po_qty
        
        status = "✅ Matched" if wo_qty == po_qty else "❌ Mismatch"
        
        comparison_data.append({
            'Size': size,
            'WO Quantity': int(wo_qty) if wo_qty == int(wo_qty) else wo_qty,
            'PO Quantity': int(po_qty) if po_qty == int(po_qty) else po_qty,
            'Difference': int(difference) if difference == int(difference) else difference,
            'Status': status
        })
    
    # Display comparison table
    comparison_df = pd.DataFrame(comparison_data)
    st.dataframe(comparison_df, use_container_width=True, hide_index=True)
    
    # Summary metrics
    total_wo = sum(item['WO Quantity'] for item in comparison_data)
    total_po = sum(item['PO Quantity'] for item in comparison_data)
    total_diff = total_wo - total_po
    matched_sizes = len([item for item in comparison_data if item['Status'] == '✅ Matched'])
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total WO Qty", total_wo)
    with col2:
        st.metric("Total PO Qty", total_po)
    with col3:
        st.metric("Difference", total_diff)
    with col4:
        st.metric("Matched Sizes", matched_sizes)
    
    # --- FIX IS HERE ---
    # Create a unique key for the download button using the PO number
    download_key = f"download_size_comparison_{matched_po_number}"
    
    # Download button
    csv = comparison_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="⬇️ Download Size Comparison as CSV",
        data=csv,
        file_name=f"wo_po_size_comparison_{matched_po_number}.csv",
        mime="text/csv",
        key=download_key  # Use the unique key here
    )

# Keep this for backward compatibility
def display_size_comparison_for_matched_po(wo_items: List[Dict], po_list: List[Dict[str, Any]], matched_po_number: str):
    """
    Legacy function - redirects to new implementation.
    """
    matched_po = None
    for po in po_list:
        if po.get('po_number') == matched_po_number:
            matched_po = po
            break
    
    if matched_po:
        # Pass the matched_po_number to the updated function
        display_size_comparison(matched_po, wo_items, matched_po_number)
    else:
        st.error(f"❌ Could not find PO with number: {matched_po_number}")