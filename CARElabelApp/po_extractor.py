"""
PO Data Extractor Module
==================

This module extracts structured data from merged Purchase Order (PO) PDFs.
Now includes advanced parsing for Color/Size/Destination 3rd-line fields:
    e.g. 
    1️⃣ "L C1-430422-QED/36013805/12 25 / XMill Date : 22-10-25"
        → Size = L | Color Code = C1 | VSD = 430422-QED | Factory ID = 36013805
    2️⃣ "421004-RAS/36013805/11 25 XS XMill Date : 27-10-25"
        → Color Code = 421004 | VSD = RAS | Factory ID = 36013805 | Size = XS
    3️⃣ "Color/Size/Destination : 421004-RAS-36013805-11/25 XL"
        → VSD = 421004-RAS | Factory ID = 36013805 | Size = XL
"""

import streamlit as st
import re
import pdfplumber
import pandas as pd
from typing import List, Dict, Any, Optional

# =============================================================================
# MAIN EXTRACTION FUNCTIONS
# =============================================================================

def extract_merged_po_details(pdf_file) -> List[Dict[str, Any]]:
    po_list = []
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page_num in range(1, len(pdf.pages)):
                page = pdf.pages[page_num]
                text = page.extract_text() or ""
                if "Purchase Order" in text and "PO No." in text:
                    po_number = extract_po_number(text)
                    supplier_info = extract_supplier_info(text)
                    items = extract_po_items_enhanced(text)
                    po_details = extract_additional_po_details(text)
                    
                    # REPLACED WITH SAFER VERSION:
                    total_quantity = 0
                    for item in items:
                        try:
                            if item.get('quantity') and item['quantity'] != '':
                                total_quantity += float(item['quantity'])
                        except (ValueError, TypeError):
                            # Skip invalid quantities
                            continue
                    
                    po_list.append({
                        'po_number': po_number,
                        'supplier': supplier_info['supplier'],
                        'items': items,
                        'total_quantity': int(total_quantity),  # Convert to int at end
                        **po_details
                    })
    except Exception as e:
        st.error(f"Error extracting PO details: {str(e)}")
        return []
    return po_list

def extract_delivery_location(text: str) -> str:
    """
    Extracts delivery location text from "Delivery Location" section at the end of PO.
    Extracts content under "Delivery Location" until "This is an automated PO." is found.
    """
    # In PO: Extract content under "Delivery Location" section
    lines = text.split('\n')
    delivery_location_parts = []
    capture_started = False
    
    for i, line in enumerate(lines):
        line_stripped = line.strip()
        
        # Look for "Delivery Location" marker (case insensitive)
        if 'delivery location' in line_stripped.lower():
            capture_started = True
            
            # Check if there's content on the same line after "Delivery Location"
            # Handle formats like "Delivery Location Forwarder:" or "Delivery Location:"
            if ':' in line_stripped:
                # Extract any content after the colon on the same line
                after_colon = line_stripped.split(':', 1)[1].strip()
                if after_colon and not after_colon.lower().startswith('forwarder'):
                    delivery_location_parts.append(after_colon)
            continue
        
        # If we've started capturing and found a non-empty line
        if capture_started and line_stripped:
            # Stop capturing if we hit "This is an automated PO."
            if 'this is an automated po.' in line_stripped.lower():
                break
            
            # Add the line if it's not a header or separator
            if not line_stripped.startswith('=') and not line_stripped.startswith('-'):
                delivery_location_parts.append(line_stripped)
        
        # Stop if we've captured content and hit multiple empty lines
        elif capture_started and not line_stripped:
            # Check if the next few lines are also empty (end of section)
            empty_count = 0
            for j in range(i + 1, min(i + 4, len(lines))):
                if not lines[j].strip():
                    empty_count += 1
                else:
                    break
            if empty_count >= 2:  # Multiple empty lines indicate section end
                break
    
    if delivery_location_parts:
        # Clean and join the parts
        result = ' '.join(delivery_location_parts).strip()
        # Remove any remaining "Forwarder:" text if present
        result = re.sub(r'^forwarder:\s*', '', result, flags=re.IGNORECASE)
        return result if result else "Not found"
    
    # Fallback: Try regex patterns for specific formats
    patterns = [
        # Pattern for "Delivery Location:" followed by content until "This is an automated PO."
        r"Delivery Location\s*:\s*\n(.*?)(?=This is an automated PO\.)",
        
        # Pattern for "Delivery Location Forwarder:" followed by content
        r"Delivery Location\s+Forwarder:\s*\n(.*?)(?=This is an automated PO\.)",
        
        # Simple pattern to get content after "Delivery Location:" until next major section
        r"Delivery Location[^:\n]*:?\s*\n(.*?)(?=\n\s*\n|\n.*(?:This is an automated|Please send|Brandix Apparel))"
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            result = match.group(1).strip()
            # Clean up the result
            result = ' '.join(result.split())
            result = re.sub(r'^forwarder:\s*', '', result, flags=re.IGNORECASE)
            if result and len(result) > 5:
                return result
    
    return "Not found"

def extract_additional_po_details(text: str) -> Dict[str, str]:
    details = {}
    forwarder_match = re.search(r'Delivery Location Forwarder:\s*(.*?)\n', text)
    if forwarder_match:
        details['delivery_forwarder'] = forwarder_match.group(1).strip()
    payment_terms_match = re.search(r'Payment Terms:\s*(.*?)\n', text)
    if payment_terms_match:
        details['payment_terms'] = payment_terms_match.group(1).strip()
    packaging_terms_match = re.search(r'Packaging Terms:\s*(.*?)\n', text)
    if packaging_terms_match:
        details['packaging_terms'] = packaging_terms_match.group(1).strip()
    
    # This line calls the extract_delivery_location function
    details['delivery_location'] = extract_delivery_location(text)
    return details

def display_merged_po_results(po_list: List[Dict[str, Any]]):
    if not po_list:
        st.warning("No PO details found in uploaded PDF.")
        return
    st.success(f"✅ Extracted {len(po_list)} PO(s) successfully.")
    apply_table_styles()

    # Group POs by their PO number
    grouped_pos = {}
    for po in po_list:
        po_number = po['po_number']
        if po_number not in grouped_pos:
            grouped_pos[po_number] = []
        grouped_pos[po_number].append(po)

    # Now, iterate through the grouped POs
    for i, (po_number, pos_with_same_number) in enumerate(grouped_pos.items()):
        
        # Combine all items from POs with the same number
        all_items = []
        for po in pos_with_same_number:
            all_items.extend(po.get('items', []))

        # Check if there are any items to display
        if not all_items:
            continue

        # We have items, so proceed to display
        supplier_info = pos_with_same_number[0]['supplier']
        delivery_location = pos_with_same_number[0].get('delivery_location', '')
        
        # Create a combined PO dictionary to pass to the table creation function
        combined_po = {
            'po_number': po_number,
            'supplier': supplier_info,
            'items': all_items,
            'total_quantity': sum(safe_float_conversion(item['quantity']) for item in all_items if is_number(item['quantity'])),
            'delivery_location': delivery_location
        }

        # Display PO header with PO number, supplier, and delivery location
        st.markdown(f"""
        <div class="po-card">
            <h3>PO Number: {po_number}</h3>
            <p><strong>Supplier:</strong> {supplier_info}</p>
            <p><strong>Delivery Location:</strong> {delivery_location}</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Create a single table for all items of this PO number
        po_df = create_po_table(combined_po)
        
        # Display the table and download button
        st.dataframe(po_df, use_container_width=True)
        
        # Add download button for this specific PO with a unique key
        csv = po_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            f"⬇️ Download PO {po_number} Data", 
            csv, 
            f"po_{po_number}_data.csv", 
            "text/csv",
            key=f"download_{po_number}_{i}"  # Unique key using PO number and group index
        )
        
        # Add a separator between POs (only if items were shown)
        st.markdown("---")

def extract_po_number(text: str) -> str:
    po_match = re.search(r'BFF\s+(\d+)', text)
    return po_match.group(1) if po_match else ""


def extract_supplier_info(text: str) -> Dict[str, str]:
    supplier_info = {
        'supplier': '',
        'attention': '',
        'vat_reg_no': '',
        'business_reg_no': '',
        'telephone': '',
        'email': ''
    }

    attention_match = re.search(r'Attention\s*:\s*(.*?)\n', text)
    if attention_match:
        supplier_info['attention'] = attention_match.group(1).strip()

    supplier_match = re.search(r'Supplier\s*:\s*(.*?)(?:\n|$)', text)
    if supplier_match:
        supplier_name = supplier_match.group(1).strip()
        stop_words = ['Printout Date', 'Requested By', 'Approver', 'PO No', 'Date']
        for word in stop_words:
            if word in supplier_name:
                supplier_name = supplier_name.split(word)[0].strip()
        supplier_info['supplier'] = supplier_name

    vat_match = re.search(r'VAT Reg\.?\s*No\.?\s*:\s*(.*?)\n', text)
    if vat_match:
        supplier_info['vat_reg_no'] = vat_match.group(1).strip()

    business_match = re.search(r'Business\s*Rg\.?\s*No\.?\s*:\s*(.*?)\n', text)
    if business_match:
        supplier_info['business_reg_no'] = business_match.group(1).strip()

    telephone_match = re.search(r'T\.?\s*Phone\s*/?\s*Telefax\s*:\s*(.*?)\n', text)
    if telephone_match:
        supplier_info['telephone'] = telephone_match.group(1).strip()

    email_match = re.search(r'E\s*Mail\s*:\s*(.*?)\n', text)
    if email_match:
        supplier_info['email'] = email_match.group(1).strip()

    return supplier_info

def extract_po_items_enhanced(text: str) -> List[Dict[str, str]]:
    items = []
    lines = text.split('\n')
    table_start = find_table_start(lines)
    if table_start == -1:
        return items
    i = table_start + 1
    while i < len(lines):
        line = lines[i].strip()
        if re.match(r'^\d+\s+', line):
            item = extract_item_details(lines, i)
            if item:
                items.append(item)
                if i + 2 < len(lines):
                    i += 2
                else:
                    i += 1
        i += 1
    return items


def find_table_start(lines: List[str]) -> int:
    for i, line in enumerate(lines):
        if "No Item Quantity" in line or ("No Item" in line and "Quantity" in line):
            return i
    return -1


def extract_item_details(lines: List[str], line_index: int) -> Optional[Dict[str, str]]:
    line = lines[line_index].strip()
    qty_match = re.search(r'([\d,]+\.\d+)', line)
    if not qty_match:
        return None
    quantity = qty_match.group(1).replace(',', '')
    desc_match = re.search(r'\d+\s+(.*?)\s+([\d,]+\.\d+)', line)
    description = desc_match.group(1).strip() if desc_match else ""
    
    # Check second line for TOP/T/BOTTOM/B and append to description if found
    if line_index + 1 < len(lines):
        second_line = lines[line_index + 1].strip()
        # Check if second line contains TOP, T, BOTTOM, or B
        if re.search(r'\b(TOP|T|BOTTOM|B)\b', second_line, re.IGNORECASE):
            # Add the relevant part to description
            description += " " + second_line
    
    # Remove "Pieces" word from description (case-insensitive)
    description = re.sub(r'\bPieces\b', '', description, flags=re.IGNORECASE).strip()
    
    product_code, color_code_from_product = extract_product_code(description)
    price_match = re.search(r'([\d,]+\.\d+)\s+([\d,]+\.\d+)', line)
    unit_price = price_match.group(1) if price_match else ""
    line_amount = price_match.group(2) if price_match else ""
    size_info = extract_size_and_details_from_next_line(lines, line_index)

    return {
        'description': description,
        'product_code': product_code,
        'quantity': str(quantity),  
        'unit_price': unit_price,
        'line_amount': line_amount,
        'size': size_info.get('size', ''),
        'color_code': size_info.get('color_code', ''),
        'color_code_from_product': color_code_from_product,
        'vsd': size_info.get('vsd', ''),
        'factory_id': size_info.get('factory_id', ''),
        'color_size_destination': size_info.get('full_details', ''),
        'others': size_info.get('others', ''),
        'third_line': size_info.get('third_line', ''),
        'date_of_mfr': size_info.get('date_of_mfr', '')
    }


def extract_product_code(description: str) -> tuple[str, str]:
    """
    Extracts product code and color code from description.
    Returns a tuple of (product_code, color_code)
    """
    code_match = re.search(r'LBL\.CARE_(.*?)(?:-MWC|$)', description)
    if code_match:
        full_code = code_match.group(1).strip()
        # Split the full code into product code and color code
        if '-' in full_code:
            parts = full_code.split('-', 1)
            product_code = parts[0]
            color_code = parts[1]
        else:
            product_code = full_code
            color_code = ""
        return product_code, color_code
    return "", ""


def extract_size_from_line(line: str) -> str:
    """
    Extracts size from line before "XMill Date".
    Validates that the size is one of: XS, S, M, L, XL, XXL
    """
    # Available sizes
    valid_sizes = ['XS', 'S', 'M', 'L', 'XL', 'XXL']
    
    # Look for size before "XMill Date"
    # Pattern 1: Size at the end before XMill
    size_match1 = re.search(r'\b(' + '|'.join(valid_sizes) + r')\b(?=\s+XMill)', line)
    # Pattern 2: Size with space before XMill
    size_match2 = re.search(r'\b(' + '|'.join(valid_sizes) + r')\b\s+(?=XMill)', line)
    # Pattern 3: Size anywhere in the line
    size_match3 = re.search(r'\b(' + '|'.join(valid_sizes) + r')\b', line)
    
    if size_match1:
        return size_match1.group(1)
    elif size_match2:
        return size_match2.group(1)
    elif size_match3:
        return size_match3.group(1)
    
    return ""


def extract_size_from_third_line(third_line: str) -> str:
    """
    Extracts size specifically from third line.
    Validates that the size is one of: XS, S, M, L, XL, XXL
    """
    # Available sizes
    valid_sizes = ['XS', 'S', 'M', 'L', 'XL', 'XXL']
    
    # Pattern 1: Size at the beginning of line
    size_match1 = re.search(r'^\b(' + '|'.join(valid_sizes) + r')\b', third_line)
    # Pattern 2: Size anywhere in the line
    size_match2 = re.search(r'\b(' + '|'.join(valid_sizes) + r')\b', third_line)
    
    if size_match1:
        return size_match1.group(1)
    elif size_match2:
        return size_match2.group(1)
    
    return ""


def extract_factory_id_from_third_line(third_line: str) -> str:
    """
    Extracts 8-digit factory ID from third line.
    """
    # Look for any 8-digit number in the third line
    factory_match = re.search(r'\b(\d{8})\b', third_line)
    if factory_match:
        return factory_match.group(1)
    
    return ""


def extract_vsd_from_third_line(third_line: str) -> str:
    """
    Extracts VSD code from third line.
    Pattern: 6 digits, hyphen, 3 capital letters (e.g., 421015-QMW)
    """
    # Look for pattern: 6 digits, hyphen, 3 capital letters
    vsd_match = re.search(r'\b(\d{6}-[A-Z]{3})\b', third_line)
    if vsd_match:
        return vsd_match.group(1)
    
    return ""


def extract_date_of_mfr_from_third_line(third_line: str) -> str:
    """
    Extracts Date of MFR from third line.
    Handles patterns like "11 25" or "11-25" and converts to "11/25"
    """
    # Look for pattern: 2 digits, space or hyphen, 2 digits
    date_match = re.search(r'\b(\d{2})\s*[-/]\s*(\d{2})\b', third_line)
    if date_match:
        return f"{date_match.group(1)}/{date_match.group(2)}"
    
    return ""


def extract_color_code_from_third_line(third_line: str) -> str:
    """
    Extracts color code from third line.
    Looks for patterns like C1, C307, or C/1.
    """
    # Look for color code patterns
    color_match = re.search(r'\b(C1|C307|C/1)\b', third_line)
    if color_match:
        return color_match.group(1)
    
    return ""


def extract_size_and_details_from_next_line(lines: List[str], line_index: int) -> Dict[str, str]:
    result = {
        'size': '',
        'full_details': '',
        'others': '',
        'third_line': '',
        'color_code': '',
        'vsd': '',
        'factory_id': '',
        'date_of_mfr': ''
    }

    if line_index + 1 >= len(lines):
        return result

    next_line = lines[line_index + 1].strip()
    if "Color/Size/Destination" in next_line:
        match = re.search(r'Color/Size/Destination\s*:\s*(.*)', next_line)
        if match:
            result['full_details'] = match.group(1).strip()
            
            # Extract VSD (6 digits, hyphen, 3 letters)
            # Pattern: 6 digits, hyphen, 3 letters
            vsd_match = re.search(r'(\d{6}-[A-Z]{3})', result['full_details'])
            if vsd_match:
                result['vsd'] = vsd_match.group(1)
            
            # Extract factory ID (8 digits after the VSD code)
            # Look for factory ID after VSD code
            if result['vsd']:
                factory_match = re.search(r'(?:' + re.escape(result['vsd']) + r')[-/](\d{8})', result['full_details'])
                if factory_match:
                    result['factory_id'] = factory_match.group(1)
            
            # Extract size using the new function
            result['size'] = extract_size_from_line(result['full_details'])

    # Handle third line (existing logic for color code and factory ID)
    if line_index + 2 < len(lines):
        third_line = lines[line_index + 2].strip()
        # Convert all slashes to hyphens in third line
        third_line = third_line.replace('/', '-')
        result['third_line'] = third_line

        # Extract size specifically from third line
        size_from_third = extract_size_from_third_line(third_line)
        if size_from_third:
            result['size'] = size_from_third

        # Extract VSD specifically from third line
        vsd_from_third = extract_vsd_from_third_line(third_line)
        if vsd_from_third:
            result['vsd'] = vsd_from_third

        # Extract factory ID from third line using new function
        factory_id_from_third = extract_factory_id_from_third_line(third_line)
        if factory_id_from_third:
            result['factory_id'] = factory_id_from_third

        # Extract Date of MFR from third line
        date_of_mfr = extract_date_of_mfr_from_third_line(third_line)
        if date_of_mfr:
            result['date_of_mfr'] = date_of_mfr
            
        # Extract color code from third line
        color_code_from_third = extract_color_code_from_third_line(third_line)
        if color_code_from_third:
            result['color_code'] = color_code_from_third

        # Pattern 1 → size first
        pattern1 = r'(?P<size>[A-Z]{1,3})\s+(?P<color>[A-Z0-9-]+)-(?P<factory>\d+)'
        # Pattern 2 → size last
        pattern2 = r'(?P<color>[A-Z0-9]+)-(?P<factory>\d+).*?\b(?P<size>[A-Z]{1,3})\b(?=\s*XMill|$)'

        match = re.search(pattern1, third_line) or re.search(pattern2, third_line)
        if match:
            # Only use size from third line if we don't already have one from Color/Size/Destination line
            if not result['size']:
                result['size'] = match.group('size')
            # Only use color code from third line if we don't already have one from Color/Size/Destination line
            if not result['color_code']:
                result['color_code'] = match.group('color')
            # Only use factory ID from third line if we don't already have one from Color/Size/Destination line
            if not result['factory_id']:
                result['factory_id'] = match.group('factory')

    return result

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def is_number(s):
    try:
        float(s)
        return True
    except (ValueError, TypeError):
        return False

def safe_float_conversion(value):
    try:
        return float(str(value).replace(',', ''))
    except (ValueError, TypeError):
        return 0.0

def safe_int_conversion(value):
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return 0

# ✅ NEW FUNCTION — Extract color code from Third Line column in table
def extract_color_code_from_table_third_line(third_line_text: str) -> str:
    """
    Reads the Third Line column from the table and finds color code after size.
    Looks for patterns: C1, C3007, C/1 after sizes (XS, S, M, L, XL, XXL).
    """
    if not third_line_text or pd.isna(third_line_text):
        return ""
    
    # Available sizes - order matters! XXL and XL must come before L
    valid_sizes = ['XXL', 'XL', 'XS', 'L', 'M', 'S']
    
    # First, find the size in the third line
    size_found = ""
    for size in valid_sizes:
        if re.search(r'\b' + size + r'\b', third_line_text):
            size_found = size
            break
    
    if not size_found:
        return ""
    
    # Now look for color code AFTER the size
    # Pattern: size followed by whitespace, then C + digits OR C/ + digits
    pattern = re.escape(size_found) + r'\s+(C\d+|C/\d+)'
    color_match = re.search(pattern, third_line_text)
    
    if color_match:
        return color_match.group(1)
    
    return ""

# =============================================================================
# NEW HELPER FUNCTION FOR CARE INSTRUCTIONS
# =============================================================================

def extract_care_instructions(description: str) -> str:
    """
    Extracts care instruction codes (e.g., MWC015, HWC123) from a description string.
    The pattern is MW or HW, followed by 'C', followed by digits.
    Returns the first match found, or an empty string if no match.
    """
    if not description or pd.isna(description):
        return ""
    
    # Regex to find patterns like MWC015 or HWC123
    # \b ensures we match whole words to avoid partial matches.
    # (MW|HW) matches either 'MW' or 'HW'.
    # C matches the literal character 'C'.
    # \d+ matches one or more digits.
    pattern = r'\b(MW|HW)C\d+\b'
    
    match = re.search(pattern, description, re.IGNORECASE)
    
    if match:
        return match.group(0) # group(0) returns the entire matched string
    
    return ""

# =============================================================================
# DATA PROCESSING AND DISPLAY
# =============================================================================

def create_detailed_table(po_list: List[Dict[str, Any]]) -> pd.DataFrame:
    table_data = []
    for po in po_list:
        for item in po.get('items', []):
            table_data.append({
                'PO Number': po['po_number'],
                'Supplier': po['supplier'],
                'Delivery Location': po.get('delivery_location', ''),
                'Description': item['description'],  # Kept for processing Care Instructions
                'Product Code': item.get('product_code', ''),
                'Size': item.get('size', ''),
                'Color Code': item.get('color_code_from_product', ''),
                'VSD': item.get('vsd', ''),
                'Factory ID': item.get('factory_id', ''),
                'Date of MFR': item.get('date_of_mfr', ''),
                'Quantity': safe_int_conversion(safe_float_conversion(item['quantity'])),
                'Line Amount': item['line_amount'],
                'Third Line': item.get('third_line', ''),  # Kept for processing Color Code
                'PO Total Quantity': safe_int_conversion(safe_float_conversion(po['total_quantity']))
            })
    
    df = pd.DataFrame(table_data)

    # ✅ FIX: Check if the DataFrame is empty before processing columns
    if df.empty:
        return df

    # ✅ POST-PROCESS: Extract color code from Third Line column if Color Code is empty
    if 'Third Line' in df.columns and 'Color Code' in df.columns:
        for idx, row in df.iterrows():
            if not row['Color Code'] or row['Color Code'] == '':
                # Try to extract color code from Third Line column
                color_from_third = extract_color_code_from_table_third_line(row['Third Line'])
                if color_from_third:
                    df.at[idx, 'Color Code'] = color_from_third
    
    # ✅ Add Care Instructions column before the Quantity column
    # 1. Calculate the care instructions for each row
    care_instructions_series = df['Description'].apply(extract_care_instructions)
    
    # 2. Find the integer location of the 'Quantity' column
    try:
        quantity_col_index = df.columns.get_loc('Quantity')
        
        # 3. Insert the new column at the desired position
        df.insert(quantity_col_index, 'Care Instructions', care_instructions_series)
    except KeyError:
        st.warning("Could not find 'Quantity' column to insert 'Care Instructions' before it. Adding to the end instead.")
        df['Care Instructions'] = care_instructions_series

    # ✅ HIDE COLUMNS: Drop 'Description' and 'Third Line' from the final DataFrame
    # We kept them until now because they were needed for processing.
    columns_to_drop = ['Description', 'Third Line']
    df.drop(columns=columns_to_drop, inplace=True, errors='ignore') # errors='ignore' prevents crashes if columns are already gone
    
    return df

# =============================================================================
# NEW FUNCTION TO CREATE TABLE FOR A SINGLE PO
# =============================================================================

def create_po_table(po: Dict[str, Any]) -> pd.DataFrame:
    """
    Creates a DataFrame for a single PO's items.
    """
    # Create a list with just this PO
    po_list = [po]
    
    # Use the existing create_detailed_table function
    df = create_detailed_table(po_list)
    
    # Remove PO Number and Supplier columns since they're displayed in the header
    if 'PO Number' in df.columns:
        df.drop(columns=['PO Number'], inplace=True)
    if 'Supplier' in df.columns:
        df.drop(columns=['Supplier'], inplace=True)
    
    return df

# =============================================================================
# STREAMLIT DISPLAY FUNCTIONS
# =============================================================================

def apply_table_styles():
    st.markdown("""
    <style>
    .dataframe th { text-align: left !important; font-weight: bold; padding: 10px !important; background-color: #f8f9fa !important; }
    .dataframe td { padding: 10px !important; text-align: left !important; }
    .po-card { background: white; border-radius: 8px; padding: 20px; margin: 20px 0; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    </style>
    """, unsafe_allow_html=True)