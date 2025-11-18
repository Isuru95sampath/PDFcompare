import streamlit as st
import re
import pdfplumber
import pandas as pd
from typing import List, Dict, Any, Optional


# =============================================================================
# MAIN EXTRACTION FUNCTIONS
# =============================================================================
def extract_po_numbers_from_email_body(pdf_file) -> List[str]:
    """
    Extracts PO numbers ONLY from the "Subject:" line in the email.
    Specifically handles formats like: "PO 5791097/ 5791121 /5791125 / 5791126 / 5791138 / 5791133 (N51)"
    Returns a list of PO numbers found in the subject line.
    """
    try:
        with pdfplumber.open(pdf_file) as pdf:
            # Search through all pages for email details
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                
                # Look for the "Email Details" section or the "Subject:" line directly
                if "Email Details" in text or "Subject:" in text:
                    # Look for Subject line directly
                    subject_match = re.search(r'Subject:\s*(.*?)(?:\n|Factory\s*Code:|COO:)', text, re.IGNORECASE | re.DOTALL)
                    if subject_match:
                        subject = subject_match.group(1).strip()
                        
                        print(f"DEBUG: Found subject line: {subject}") # This will help you debug
                        
                        # First, remove any text in parentheses at the end (like "(N51)")
                        subject_cleaned = re.sub(r'\s*\([^)]*\)\s*$', '', subject)
                        
                        print(f"DEBUG: Cleaned subject line: {subject_cleaned}") # This will help you debug
                        
                        # --- THIS IS THE KEY CHANGE ---
                        # This is the simplest and most robust approach:
                        # Find all 6+ digit numbers in the subject line.
                        # We assume the subject line primarily contains PO numbers.
                        # \b ensures we match whole numbers, not parts of longer numbers.
                        all_po_numbers = re.findall(r'\b(\d{6,})\b', subject_cleaned)
                        
                        print(f"DEBUG: PO numbers found by regex: {all_po_numbers}") # This will help you debug
                        
                        # Remove duplicates while preserving order
                        seen = set()
                        unique_po_numbers = []
                        for po in all_po_numbers:
                            if po not in seen:
                                seen.add(po)
                                unique_po_numbers.append(po)
                        
                        print(f"DEBUG: Final unique PO numbers: {unique_po_numbers}") # This will help you debug
                        
                        if unique_po_numbers:
                            return unique_po_numbers
    
    except Exception as e:
        print(f"Error extracting email PO numbers: {e}")
    
    return []
    

def extract_email_body_data(pdf_file) -> Optional[pd.DataFrame]:
    """
    Extracts semi-structured data from the first few pages of the PDF.
    Handles formats like:
    1. Vertical format with headers like "Style", "Color", "Care Detail"
    2. Mixed format with headers and values in different arrangements
    
    Returns a pandas DataFrame or None if no data is found.
    """
    try:
        with pdfplumber.open(pdf_file) as pdf:
            # We'll check the first 3 pages for this data
            max_pages = min(3, len(pdf.pages))
            
            # Collect all text from the first few pages
            all_text = ""
            for page_num in range(max_pages):
                page = pdf.pages[page_num]
                text = page.extract_text() or ""
                all_text += text + "\n"
            
            # Define the headers we're looking for
            headers = ["Style", "Color", "Care Detail", "PO", "STYLE", "ITEM_DESCRIPTION", 
                      "Garment", "REQUESTED", "description", "Care code + content"]
            
            # Split the text into lines
            lines = all_text.split('\n')
            
            # Initialize a dictionary to store the data
            data = {}
            current_header = None
            values = []
            
            # Process each line
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # Check if this line is a header
                is_header = False
                for header in headers:
                    if line == header or line.startswith(header):
                        # Save the previous header's data if exists
                        if current_header and values:
                            data[current_header] = values
                        
                        # Start a new header
                        current_header = header
                        values = []
                        is_header = True
                        break
                
                # If not a header, add as a value to the current header
                if not is_header and current_header:
                    values.append(line)
            
            # Don't forget the last header
            if current_header and values:
                data[current_header] = values
            
            # If we found data, convert to DataFrame
            if data:
                # Find the maximum number of values for any header
                max_values = max(len(values) for values in data.values())
                
                # Create a dictionary where each key is a header and each value is a list
                # with length equal to max_values (filling with empty strings if needed)
                df_data = {}
                for header, values in data.items():
                    df_data[header] = values + [""] * (max_values - len(values))
                
                # Create the DataFrame
                df = pd.DataFrame(df_data)
                
                return df
                
    except Exception as e:
        print(f"Error extracting email body data: {e}")
    
    return None

def extract_merged_po_details(pdf_file) -> List[Dict[str, Any]]:
    po_list = []
    try:
        # Extract all PO numbers from email body first
        email_po_numbers = extract_po_numbers_from_email_body(pdf_file)
        
        with pdfplumber.open(pdf_file) as pdf:
            # Process ALL pages for detailed extraction (including first page)
            for page_num in range(len(pdf.pages)):  # Start from 0 to include first page
                page = pdf.pages[page_num]
                text = page.extract_text() or ""
                
                # More flexible PO detection - check for multiple patterns
                is_po_page = (
                    ("Purchase Order" in text and "PO No." in text) or
                    ("BFF" in text and re.search(r'BFF\s+\d+', text)) or
                    ("PO No." in text and re.search(r'PO\s*No\.?\s*:?\s*\d+', text))
                )
                
                if is_po_page:
                    # Extract PO number using multiple methods
                    po_number = extract_po_number(text)
                    if not po_number:
                        # Try alternative extraction
                        alt_po_match = re.search(r'PO\s*No\.?\s*:?\s*(\d+)', text)
                        if alt_po_match:
                            po_number = alt_po_match.group(1)
                    
                    if po_number:  # Only process if we found a PO number
                        supplier_info = extract_supplier_info(text)
                        items = extract_po_items_enhanced(text)
                        
                        # 🔥 NEW: Consolidate duplicate sizes
                        items = consolidate_duplicate_sizes(items)
                        
                        po_details = extract_additional_po_details(text)
                        
                        # Calculate total quantity safely
                        total_quantity = 0
                        for item in items:
                            try:
                                if item.get('quantity') and item['quantity'] != '':
                                    total_quantity += float(item['quantity'])
                            except (ValueError, TypeError):
                                # Skip invalid quantities
                                continue
                        
                        # Check if this PO number matches any from the email
                        matching_email_po = ""
                        for email_po in email_po_numbers:
                            if po_number == email_po:
                                matching_email_po = email_po
                                break
                        
                        po_list.append({
                            'po_number': po_number,
                            'email_po_number': matching_email_po,  # Add the matching email PO number
                            'supplier': supplier_info['supplier'],
                            'items': items,
                            'total_quantity': int(total_quantity),  # Convert to int at end
                            'page_number': page_num + 1,  # Add page number for debugging
                            **po_details
                        })
                        
                        print(f"Extracted PO {po_number} from page {page_num + 1}")
    
    except Exception as e:
        st.error(f"Error extracting PO details: {str(e)}")
        return []
    
    print(f"Total POs extracted: {len(po_list)}")
    return po_list

def consolidate_duplicate_sizes(po_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Consolidates duplicate sizes in PO items by grouping by size and color code.
    Only combines quantities if both size AND color code match.
    """
    if not po_items:
        return []
    
    # Dictionary to store consolidated items
    consolidated_items = {}
    
    for item in po_items:
        # Get size and color code, normalize them for comparison
        size = str(item.get('size', '')).strip().upper()
        color_code = str(item.get('color_code', '')).strip().upper()
        
        # Create a key combining size and color code
        # If color code is missing, we still group by size alone
        key = (size, color_code) if color_code else (size,)
        
        if key in consolidated_items:
            # Item with same size (and color code) exists - add quantities
            existing_item = consolidated_items[key]
            
            try:
                # Get existing quantity
                existing_qty = float(str(existing_item.get('quantity', '0')).replace(',', ''))
                
                # Get new quantity
                new_qty = float(str(item.get('quantity', '0')).replace(',', ''))
                
                # Sum the quantities
                total_qty = existing_qty + new_qty
                
                # Update the quantity in the consolidated item
                existing_item['quantity'] = str(total_qty)
                
                # Update line amount if unit price is present
                if existing_item.get('unit_price'):
                    try:
                        unit_price = float(str(existing_item['unit_price']).replace(',', ''))
                        existing_item['line_amount'] = str(round(total_qty * unit_price, 2))
                    except (ValueError, TypeError):
                        pass
                
            except (ValueError, TypeError):
                # If quantity conversion fails, keep the original
                pass
        else:
            # New size/color combination - add as is
            consolidated_items[key] = item.copy()
    
    # Convert back to list
    result = list(consolidated_items.values())
    
    # Log the consolidation results
    original_count = len(po_items)
    consolidated_count = len(result)
    
    if original_count > consolidated_count:
        st.info(f"📊 Consolidated {original_count} PO items to {consolidated_count} items by combining duplicate sizes and color codes.")
    
    return result

# Add this helper function to display debugging information in your UI

def display_email_po_debug_info(pdf_file):
    """
    Displays debugging information about email PO number extraction from subject line only
    """
    try:
        with pdfplumber.open(pdf_file) as pdf:
            st.write("🔍 Debug: First page text preview:")
            first_page_text = pdf.pages[0].extract_text() or ""
            # Show first 1000 characters
            st.text_area("First 1000 characters:", first_page_text[:1000], height=200)
            
            # Look for Subject line
            subject_match = re.search(r'Subject:\s*(.*?)(?:\n|Factory\s*Code:|COO:)', first_page_text, re.IGNORECASE | re.DOTALL)
            if subject_match:
                subject = subject_match.group(1).strip()
                st.write(f"✅ Found Subject: {subject}")
                
                # Show extracted PO numbers from subject only
                email_po_numbers = extract_po_numbers_from_email_body(pdf_file)
                if email_po_numbers:
                    st.write(f"📧 PO Numbers from subject line: {email_po_numbers}")
                    st.write(f"📊 Total PO numbers in subject: {len(email_po_numbers)}")
                else:
                    st.write("❌ No PO numbers found in subject line")
            else:
                st.write("❌ No Subject line found")
                
            # Show all PDF PO numbers
            pdf_po_numbers = []
            with pdfplumber.open(pdf_file) as pdf_inner:
                for page_num in range(1, len(pdf_inner.pages)):
                    page = pdf_inner.pages[page_num]
                    text = page.extract_text() or ""
                    if "Purchase Order" in text and "PO No." in text:
                        po_number = extract_po_number(text)
                        if po_number:
                            pdf_po_numbers.append(po_number)
            
            st.write(f"📄 PDF PO Numbers: {pdf_po_numbers}")
                
    except Exception as e:
        st.error(f"Debug error: {e}")

def extract_delivery_location(text: str) -> str:
    """
    Extracts delivery location text from "Delivery Location" section at end of PO.
    Gets the warehouse code from "Delivery Location Forwarder:" line and 
    the full address that appears after "This is an automated PO" message.
    """
    location_parts = []
    
    # Step 1: Get the warehouse/location code from "Delivery Location Forwarder:" line
    delivery_forwarder_pattern = r'Delivery\s+Location\s+Forwarder\s*:\s*\n\s*([^\n]+)'
    forwarder_match = re.search(delivery_forwarder_pattern, text, re.IGNORECASE)
    
    if forwarder_match:
        location_code = forwarder_match.group(1).strip()
        if location_code and location_code.lower() != 'forwarder':
            location_parts.append(location_code)
    
    # Step 2: Find the full address after "This is an automated PO"
    # Split text into lines for easier processing
    lines = text.split('\n')
    
    # Find the index where "This is an automated PO" appears
    automated_line_idx = -1
    for i, line in enumerate(lines):
        if 'this is an automated po' in line.lower():
            automated_line_idx = i
            break
    
    if automated_line_idx != -1:
        # Start collecting address lines after the automated PO message
        # Look for the actual address lines (typically 4-5 lines)
        address_lines = []
        
        # Process lines after the automated message
        for i in range(automated_line_idx + 1, len(lines)):
            line = lines[i].strip()
            
            # Skip empty lines
            if not line:
                continue
            
            # Stop conditions - these indicate we've passed the address section
            stop_keywords = [
                'please send', 'no metal pins', 'po no/s must',
                'goods delivered', 'goods should be'
            ]
            
            if any(keyword in line.lower() for keyword in stop_keywords):
                break
            
            # Skip page markers and headers
            if any(skip in line.lower() for skip in ['page 2(2)', 'page 1(2)', 'purchase order', 'po no.', 'brandix apparel (pvt) ltd']):
                # But only skip "Brandix Apparel (Pvt) Ltd" if it has "Page" or is the header format
                if 'brandix apparel (pvt) ltd' in line.lower() and ('page' not in line.lower() and 'colombo' not in line.lower()):
                    # This is likely part of the address, not the header
                    pass
                else:
                    continue
            
            # Skip lines with strange characters or patterns that aren't addresses
            if re.match(r'^[\d\s,\.]+$', line):  # Just numbers, spaces, commas
                continue
            
            # Add valid address lines
            # These should be actual address components
            if len(line) > 2 and not line.startswith('*'):
                address_lines.append(line)
            
            # Typically address is 4-5 lines, so stop after collecting enough
            if len(address_lines) >= 5:
                break
        
        # Add the collected address lines to location parts
        location_parts.extend(address_lines)
    
    # Combine all parts into one string
    if location_parts:
        result = ' '.join(location_parts).strip()
        return result if result else "Not found"
    
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

    # Group POs by their PO number (in case of duplicates)
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
        page_number = pos_with_same_number[0].get('page_number', 'N/A')
        
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
            <h3>PO Number: {po_number} (Page {page_number})</h3>
            <p><strong>Supplier:</strong> {supplier_info}</p>
            <p><strong>Delivery Location:</strong> {delivery_location}</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Create a table for all items of this PO number
        po_df = create_detailed_table([combined_po])  # Pass as a list with one item
        
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
        
        # Create a separate table for sizes and quantities
        st.markdown("### Size and Quantity Breakdown")
        
        # Prepare data for size/quantity table
        size_data = {}
        for item in all_items:
            size = item.get('size', '').strip()
            if not size:
                continue
                
            quantity = 0
            if item.get('quantity') and item['quantity'] != '':
                try:
                    quantity = float(item['quantity'])
                except (ValueError, TypeError):
                    quantity = 0
            
            if size in size_data:
                size_data[size] += quantity
            else:
                size_data[size] = quantity
        
        # Create DataFrame for size/quantity table
        if size_data:
            # Convert to list of dicts for easier sorting
            size_list = [{'Size': size, 'Quantity': qty} for size, qty in size_data.items()]
            
            # Sort by size (standard order)
            size_order = {'XS': 0, 'S': 1, 'M': 2, 'L': 3, 'XL': 4, 'XXL': 5}
            size_list.sort(key=lambda x: size_order.get(x['Size'].upper(), 99))
            
            # Create DataFrame from the sorted list
            size_df = pd.DataFrame(size_list)
            
            # Display the size/quantity table
            st.dataframe(size_df, use_container_width=True)
            
            # Add download button for size/quantity data
            size_csv = size_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                f"⬇️ Download Size/Quantity Data", 
                size_csv, 
                f"po_{po_number}_size_quantity.csv", 
                "text/csv",
                key=f"download_size_{po_number}_{i}"
            )
            
            # 🔥 NEW: Store size breakdown data in session state for comparison
            # Convert to dictionary format for easier comparison
            po_size_dict = {}
            for item in size_list:
                po_size_dict[item['Size']] = item['Quantity']
            
            # Store in session state with the key that comparison function expects
            session_state_key = f"po_size_breakdown_{po_number}"
            st.session_state[session_state_key] = po_size_dict

            # Debug info to confirm storage
            if st.checkbox("Show Debug Info", key=f"debug_{po_number}"):
                st.write(f"🔍 Debug: Stored size breakdown for PO {po_number} in session state with key: {session_state_key}")
                st.write(f"🔍 Debug: Size breakdown data: {po_size_dict}")
            
        else:
            st.warning("No size data available for this PO")
        
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


# In po_extractor.py, find and replace the extract_item_details function with this updated version:

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

    # Reconstruct the full text of the relevant lines to search for fields
    full_text_for_search = "\n".join(lines[max(0, line_index - 2) : line_index + 4])

    return {
        'description': description,
        'product_code': product_code,
        'quantity': str(quantity),  
        'unit_price': unit_price,
        'line_amount': line_amount,
        'size': size_info.get('size', ''),
        'color_code': size_info.get('color_code', ''), # This is the primary color code extraction
        'color_code_from_product': color_code_from_product,
        'vsd': size_info.get('vsd', '') or extract_vsd_from_po(full_text_for_search), # Fallback to broader search
        'factory_id': size_info.get('factory_id', ''),
        'color_size_destination': size_info.get('full_details', ''),
        'others': size_info.get('others', ''),
        'third_line': size_info.get('third_line', ''),
        'date_of_mfr': size_info.get('date_of_mfr', ''),
        'silhouette': extract_silhouette_from_po(full_text_for_search),
        'care_instruction': extract_care_instruction_from_po(full_text_for_search)
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


def extract_color_code_from_text(text: str) -> str:
    """
    Extracts general color code patterns from any text string.
    Looks for patterns like C1, C307, C/1, etc.
    This is a robust, centralized function for color code extraction.
    """
    if not text:
        return ""
    
    # Look for general color code patterns: C followed by digits, or C/ followed by digits
    # \b ensures we match whole words to avoid partial matches.
    color_match = re.search(r'\b(C\d+|C/\d+)\b', text)
    if color_match:
        return color_match.group(1)
    
    return ""


def extract_color_code_from_third_line(third_line: str) -> str:
    """
    Extracts color code from third line.
    Updated to use the more robust extract_color_code_from_text function.
    """
    return extract_color_code_from_text(third_line)


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
            
        # Extract color code from third line using the new robust function
        color_code_from_third = extract_color_code_from_text(third_line)
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

# ✅ UPDATED FUNCTION — Extract color code from Third Line column in table
def extract_color_code_from_table_third_line(third_line_text: str) -> str:
    """
    Reads the Third Line column from the table and finds color code.
    Updated to use the more robust extract_color_code_from_text function.
    """
    if not third_line_text or pd.isna(third_line_text):
        return ""
    
    return extract_color_code_from_text(str(third_line_text))

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

# In po_extractor.py

def create_detailed_table(po_list: List[Dict[str, Any]]) -> pd.DataFrame:
    table_data = []
    for po in po_list:
        for item in po.get('items', []):
            table_data.append({
                'PO Number': po['po_number'],
                'Email PO Number': po.get('email_po_number', ''),
                'Supplier': po['supplier'],
                'Delivery Location': po.get('delivery_location', ''),
                'Description': item['description'],  # Kept for processing Care Instructions
                'Product Code': item.get('product_code', ''),
                'Size': item.get('size', ''),
                'Color Code': item.get('color_code_from_product', ''),
                'VSD': item.get('vsd', ''),
                'Silhouette': item.get('silhouette', ''),  # 🔥 ADD THIS LINE
                'Factory ID': item.get('factory_id', ''),
                'Date of MFR': item.get('date_of_mfr', ''),
                'Quantity': safe_int_conversion(safe_float_conversion(item['quantity'])),
                'Line Amount': item['line_amount'],
                'Third Line': item.get('third_line', ''),  # Kept for processing Color Code
                'PO Total Quantity': safe_int_conversion(safe_float_conversion(po['total_quantity']))
            })
    
    df = pd.DataFrame(table_data)

    # ... (rest of the function, including adding Care Instructions column) ...

    # ✅ HIDE COLUMNS: Drop 'Description' and 'Third Line' from the final DataFrame
    # We keep 'Silhouette' now.
    columns_to_drop = ['Description', 'Third Line']
    df.drop(columns=columns_to_drop, inplace=True, errors='ignore') # errors='ignore' prevents crashes if columns are already gone
    
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


# Add these new functions to po_extractor.py

def extract_silhouette_from_po(text: str) -> str:
    """
    Extracts Silhouette from PO text.
    Looks for common patterns like "Silhouette:" or "Garment Description:".
    """
    # Pattern 1: Look for "Silhouette:" followed by the value
    silhouette_match = re.search(r'Silhouette\s*:\s*([^\n]+)', text, re.IGNORECASE)
    if silhouette_match:
        return silhouette_match.group(1).strip()
    
    # Pattern 2: Look for "Description:" and try to extract a common silhouette name
    desc_match = re.search(r'Description\s*:\s*([^\n]+)', text, re.IGNORECASE)
    if desc_match:
        desc = desc_match.group(1).strip()
        # Common silhouette names to look for in the description
        common_silhouettes = ["BRALETTE", "BRIEF", "HIPSTER", "THONG", "BOYSHORT", "BODYSUIT", "TANK", "CAMI", "SLEEP SHIRT"]
        for s in common_silhouettes:
            if s.upper() in desc.upper():
                return s # Return the found silhouette
    
    return "" # Return empty if no pattern matches

def extract_care_instruction_from_po(text: str) -> str:
    """
    Extracts Care Instruction code (e.g., MWC015, HWC123) from PO text.
    Uses the same pattern as the WO extractor.
    """
    # Regex to find patterns like MWC015 or HWC123
    pattern = r'\b(MW|HW)C\d+\b'
    match = re.search(pattern, text, re.IGNORECASE)
    
    if match:
        return match.group(0) # group(0) returns the entire matched string
    
    return ""

def extract_vsd_from_po(text: str) -> str:
    """
    Extracts VSD code from PO text with a more flexible pattern.
    Handles formats like "421015-QMW" or "421015 QMW".
    """
    # Pattern 1: 6 digits, optional space/hyphen, 3 capital letters
    vsd_match = re.search(r'(\d{6})[\s-]?([A-Z]{3})', text)
    if vsd_match:
        # Reconstruct with hyphen
        return f"{vsd_match.group(1)}-{vsd_match.group(2)}"
    
    return ""


def display_consolidation_summary(po_items: List[Dict[str, Any]], consolidated_items: List[Dict[str, Any]]):
    """
    Displays a summary of the consolidation process showing which items were combined.
    """
    if len(po_items) == len(consolidated_items):
        return  # No consolidation occurred
    
    st.markdown("### 📋 Consolidation Summary")
    
    # Create a summary of what was consolidated
    consolidation_summary = {}
    
    # Group original items by size and color
    for item in po_items:
        size = str(item.get('size', '')).strip().upper()
        color_code = str(item.get('color_code', '')).strip().upper()
        
        if size and color_code:
            key = f"{size} / {color_code}"
            if key not in consolidation_summary:
                consolidation_summary[key] = {
                    'size': size,
                    'color_code': color_code,
                    'original_items': [],
                    'total_quantity': 0
                }
            
            try:
                qty = float(str(item.get('quantity', '0')).replace(',', ''))
                consolidation_summary[key]['original_items'].append({
                    'description': item.get('description', ''),
                    'quantity': qty
                })
                consolidation_summary[key]['total_quantity'] += qty
            except (ValueError, TypeError):
                pass
    
    # Display the summary
    for key, data in consolidation_summary.items():
        if len(data['original_items']) > 1:  # Only show if there were duplicates
            with st.expander(f"🔗 Combined: {data['size']} / {data['color_code']} (Total: {data['total_quantity']})"):
                for i, original_item in enumerate(data['original_items'], 1):
                    st.write(f"{i}. Qty: {original_item['quantity']} - {original_item['description']}")


def create_consolidated_po_table(po: Dict[str, Any]) -> pd.DataFrame:
    """
    Creates a DataFrame for a single PO's items with special handling for consolidated items.
    Shows original descriptions in a tooltip or expandable section.
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
    
    # Add a column to indicate if this item was consolidated
    if 'original_descriptions' in po.get('items', [{}])[0]:
        df['Consolidated'] = df.apply(
            lambda row: 'Yes' if len(row.get('original_descriptions', [])) > 1 else 'No',
            axis=1
        )
    
    return df

def sort_po_items_by_size(po_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Sorts a list of PO items from smallest size to largest size.
    Handles standard sizes: XS, S, M, L, XL, XXL, XXXL.
    Items with unrecognizable sizes are placed at the end.
    """
    if not po_items:
        return []

    # Define the standard size order
    size_order = {
        'XXXL': 0, '3XL': 0,
        'XXL': 1, '2XL': 1,
        'XL': 2,
        'L': 3,
        'M': 4,
        'S': 5,
        'XS': 6,
    }

    def get_sort_key(item):
        size = str(item.get('size', '')).strip().upper()
        # Get the sort order from the dictionary, defaulting to a high number for unknown sizes
        return size_order.get(size, 99)

    # Sort the list of items using the key function
    sorted_items = sorted(po_items, key=get_sort_key)
    
    return sorted_items

def extract_garment_description_table(pdf_file) -> Optional[pd.DataFrame]:
    """
    Extracts the Garment description table from the "email body" section of the merged PO PDF.
    Returns a pandas DataFrame with the table data, or None if no table is found.
    """
    try:
        with pdfplumber.open(pdf_file) as pdf:
            # We'll check the first few pages for this data
            max_pages = min(3, len(pdf.pages))
            
            # Collect all text from the first few pages
            all_text = ""
            for page_num in range(max_pages):
                page = pdf.pages[page_num]
                text = page.extract_text() or ""
                all_text += text + "\n"
            
            # Split the text into lines
            lines = all_text.split('\n')
            
            # Look for the section that contains the Garment description table
            # This is a heuristic approach - you may need to adjust based on your actual PDF format
            table_start = -1
            table_end = -1
            
            for i, line in enumerate(lines):
                # Look for the header of the table
                if "Garment description" in line or "Garment Description" in line:
                    table_start = i
                    # Look for the end of the table (next major section or end of page)
                    for j in range(i + 1, len(lines)):
                        if lines[j].strip() == "" or "Email Details" in lines[j] or "PO #" in lines[j]:
                            table_end = j
                            break
                    
                    if table_end == -1:
                        table_end = len(lines)
                    break
            
            if table_start == -1:
                return None
            
            # Extract the table lines
            table_lines = lines[table_start:table_end]
            
            # Try to extract table data
            # This is a simplified approach - you may need to adjust based on your actual table format
            table_data = []
            headers = []
            
            # First, try to identify the headers
            if table_lines:
                # The first line might contain the headers
                header_line = table_lines[0]
                headers = [col.strip() for col in header_line.split('\t') if col.strip()]
                
                # If we don't have clear headers, try to infer them
                if len(headers) < 2:
                    # Try a different approach to identify headers
                    # Look for common header terms
                    common_headers = ["PO_NO", "PO Number", "Style", "Color", "Description", "Quantity"]
                    for header in common_headers:
                        if header in header_line:
                            headers.append(header)
                
                # Process the data rows
                for line in table_lines[1:]:
                    if line.strip():
                        # Split the line into columns
                        columns = [col.strip() for col in line.split('\t')]
                        
                        # If we don't have enough columns, try a different split method
                        if len(columns) < len(headers):
                            # Try splitting by multiple spaces
                            columns = [col.strip() for col in re.split(r'\s{2,}', line)]
                        
                        # Create a dictionary for this row
                        if len(columns) >= len(headers):
                            row_data = {headers[i]: columns[i] for i in range(len(headers))}
                            table_data.append(row_data)
            
            # Convert to DataFrame
            if table_data:
                df = pd.DataFrame(table_data)
                return df
                
    except Exception as e:
        print(f"Error extracting garment description table: {e}")
    
    return None

def filter_garment_description_by_po(garment_df, po_number):
    """
    Filters the garment description DataFrame to only include rows where PO_NO matches the given PO number.
    Returns the filtered DataFrame.
    """
    if garment_df is None or po_number is None:
        return None
    
    # Try different possible column names for PO number
    po_column = None
    for col in garment_df.columns:
        if "po" in col.lower() and "no" in col.lower():
            po_column = col
            break
    
    if po_column is None:
        # If we can't find a PO column, return the original DataFrame
        return garment_df
    
    # Filter the DataFrame
    filtered_df = garment_df[garment_df[po_column].astype(str).str.contains(po_number, na=False)]
    
    return filtered_df

def extract_email_body_item_data(pdf_path: str) -> Optional[pd.DataFrame]:
    """
    Extracts the 'PO NO' (Line Item/Color Code) and 'Garment description' columns 
    from the item table found under the 'Email Body:' section on the first page of the merged PO PDF.

    NOTE: Hardcoded column indices are used to reliably target the data due to the 
    complex, multi-line, and poorly parsed headers in the PDF.
    
    The column mapping for your specific PDF is:
    - Index 5: 'PO' column content (Line Item Code / Color Code). Mapped to 'PO NO'.
    - Index 8: 'Garment description' column content. Mapped to 'Garment description'.

    Args:
        pdf_path: The file path to the merged PO PDF.

    Returns:
        A pandas DataFrame with the requested columns, or None if extraction fails.
    """
    try:
        # 1. Open the PDF
        with pdfplumber.open(pdf_path) as pdf:
            # The item table is on the first page (index 0)
            first_page = pdf.pages[0]
            
            # 2. Extract the first table found (the item list)
            # Use default settings which work best for this specific table structure
            tables = first_page.extract_tables()
            
            if not tables or len(tables) < 1:
                return None
            
            # Select the first table data
            table_data = tables[0]
            
            if len(table_data) < 2:
                return None
            
            # The first row (index 0) is the multi-line header. Data starts from row index 1.
            data_rows = table_data[1:]

            # 3. Define the column indices based on confirmed structure:
            PO_NO_CODE_IDX = 5 
            GARMENT_DESC_IDX = 8 

            extracted_data: List[Dict[str, Any]] = []
            
            for row in data_rows:
                # Ensure the row has enough columns
                if len(row) > max(PO_NO_CODE_IDX, GARMENT_DESC_IDX):
                    
                    po_no_code = str(row[PO_NO_CODE_IDX]).strip() if row[PO_NO_CODE_IDX] else ""
                    garment_desc = str(row[GARMENT_DESC_IDX]).strip() if row[GARMENT_DESC_IDX] else ""
                    
                    # Skip rows where essential data is missing
                    if po_no_code and garment_desc:
                        extracted_data.append({
                            # Using the user's requested heading
                            "PO NO": po_no_code, 
                            "Garment description": garment_desc 
                        })
            
            # 4. Convert to DataFrame and return
            if extracted_data:
                df = pd.DataFrame(extracted_data)
                return df
            else:
                return None

    except Exception as e:
        # In a production app, you might want to log this error instead of printing
        # print(f"An error occurred during PDF extraction: {e}") 
        return None

def extract_garment_description_table(pdf_file) -> Optional[pd.DataFrame]:
    """
    Extracts the Garment description table from the "email body" section of the merged PO PDF.
    Returns a pandas DataFrame with the table data, or None if no table is found.
    """
    try:
        with pdfplumber.open(pdf_file) as pdf:
            # We'll check the first few pages for this data
            max_pages = min(3, len(pdf.pages))
            
            # Collect all text from the first few pages
            all_text = ""
            for page_num in range(max_pages):
                page = pdf.pages[page_num]
                text = page.extract_text() or ""
                all_text += text + "\n"
            
            # Split the text into lines
            lines = all_text.split('\n')
            
            # Look for the section that contains the Garment description table
            # This is a heuristic approach - you may need to adjust based on your actual PDF format
            table_start = -1
            table_end = -1
            
            for i, line in enumerate(lines):
                # Look for the header of the table
                if "Garment description" in line or "Garment Description" in line:
                    table_start = i
                    # Look for the end of the table (next major section or end of page)
                    for j in range(i + 1, len(lines)):
                        if lines[j].strip() == "" or "Email Details" in lines[j] or "PO #" in lines[j]:
                            table_end = j
                            break
                    
                    if table_end == -1:
                        table_end = len(lines)
                    break
            
            if table_start == -1:
                return None
            
            # Extract the table lines
            table_lines = lines[table_start:table_end]
            
            # Try to extract table data
            # This is a simplified approach - you may need to adjust based on your actual table format
            table_data = []
            headers = []
            
            # First, try to identify the headers
            if table_lines:
                # The first line might contain the headers
                header_line = table_lines[0]
                headers = [col.strip() for col in header_line.split('\t') if col.strip()]
                
                # If we don't have clear headers, try to infer them
                if len(headers) < 2:
                    # Try a different approach to identify headers
                    # Look for common header terms
                    common_headers = ["PO_NO", "PO Number", "Style", "Color", "Description", "Quantity"]
                    for header in common_headers:
                        if header in header_line:
                            headers.append(header)
                
                # Process the data rows
                for line in table_lines[1:]:
                    if line.strip():
                        # Split the line into columns
                        columns = [col.strip() for col in line.split('\t')]
                        
                        # If we don't have enough columns, try a different split method
                        if len(columns) < len(headers):
                            # Try splitting by multiple spaces
                            columns = [col.strip() for col in re.split(r'\s{2,}', line)]
                        
                        # Create a dictionary for this row
                        if len(columns) >= len(headers):
                            row_data = {headers[i]: columns[i] for i in range(len(headers))}
                            table_data.append(row_data)
            
            # Convert to DataFrame
            if table_data:
                df = pd.DataFrame(table_data)
                return df
                
    except Exception as e:
        print(f"Error extracting garment description table: {e}")
    
    return None

def filter_garment_description_by_po(garment_df, po_number):
    """
    Filters the garment description DataFrame to only include rows where PO_NO matches the given PO number.
    Returns the filtered DataFrame.
    """
    if garment_df is None or po_number is None:
        return None
    
    # Try different possible column names for PO number
    po_column = None
    for col in garment_df.columns:
        if "po" in col.lower() and "no" in col.lower():
            po_column = col
            break
    
    if po_column is None:
        # If we can't find a PO column, return the original DataFrame
        return garment_df
    
    # Filter the DataFrame
    filtered_df = garment_df[garment_df[po_column].astype(str).str.contains(po_number, na=False)]
    
    return filtered_df