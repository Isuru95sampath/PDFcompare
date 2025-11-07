from turtle import st
import pdfplumber
import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO
from fuzzywuzzy import fuzz


def uploaded_file_to_bytesio(uploaded_file):
    """Convert an uploaded file to a BytesIO object"""
    if uploaded_file is None:
        return None
    
    # Read the file content into memory
    file_content = uploaded_file.read()
    
    # Create a BytesIO object from the content
    bytes_io = BytesIO(file_content)
    
    # Reset the file pointer for the original uploaded file
    uploaded_file.seek(0)
    
    return bytes_io

def create_styles_pdf(styles: list) -> BytesIO:
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

def merge_pdfs(original_pdf: BytesIO, styles_pdf: BytesIO) -> BytesIO:
    pdf_out = fitz.open()
    pdf_styles = fitz.open(stream=styles_pdf.read(), filetype="pdf")
    pdf_orig = fitz.open(stream=original_pdf.read(), filetype="pdf")
    
    pdf_out.insert_pdf(pdf_styles)
    pdf_out.insert_pdf(pdf_orig)
    
    output = BytesIO()
    pdf_out.save(output)
    output.seek(0)
    return output

def merge_pdfs_with_po(styles_pdf, original_pdf, po_pdf=None):
    """
    Merge multiple PDFs: styles PDF, original PDF, and optionally PO PDF
    
    Args:
        styles_pdf: BytesIO object containing styles PDF
        original_pdf: BytesIO object containing original PDF
        po_pdf: Optional BytesIO object containing PO PDF
    
    Returns:
        BytesIO object containing merged PDF
    """
    try:
        pdf_out = fitz.open()
        
        # Add styles PDF
        if styles_pdf:
            pdf_styles = fitz.open(stream=styles_pdf.read(), filetype="pdf")
            pdf_out.insert_pdf(pdf_styles)
            pdf_styles.close()
        
        # Add original PDF
        if original_pdf:
            pdf_orig = fitz.open(stream=original_pdf.read(), filetype="pdf")
            pdf_out.insert_pdf(pdf_orig)
            pdf_orig.close()
        
        # Add PO PDF if provided
        if po_pdf:
            pdf_po = fitz.open(stream=po_pdf.read(), filetype="pdf")
            pdf_out.insert_pdf(pdf_po)
            pdf_po.close()
        
        output = BytesIO()
        pdf_out.save(output)
        output.seek(0)
        pdf_out.close()
        
        return output
        
    except Exception as e:
        st.error(f"Error merging PDFs: {e}")
        return None

def extract_style_numbers_from_po_first_page(pdf_file):
    """Extract style numbers from the first page of PO PDF"""
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            if len(pdf.pages) > 0:
                first_page_text = pdf.pages[0].extract_text() or ""
                # Look for "Extracted Style Numbers:" section
                extracted_section_match = re.search(r'Extracted Style Numbers:\s*(.+)', first_page_text, re.IGNORECASE)
                if extracted_section_match:
                    extracted_styles = re.findall(r'\b\d{8}\b', extracted_section_match.group(1))
                    return extracted_styles
                
                # Fallback: look for any 8-digit numbers
                style_numbers = re.findall(r'\b\d{8}\b', first_page_text)
                return style_numbers
        return []
    except Exception as e:
        st.error(f"Error extracting style numbers from PO: {e}")
        return []

def extract_po_number(pdf_file):
    """Extract PO Number from PO PDF"""
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            # Check first page for PO Number
            if len(pdf.pages) > 0:
                first_page_text = pdf.pages[0].extract_text() or ""
                
                # Look for "PO Number:" pattern
                po_number_match = re.search(r'PO Number:\s*(\d+)', first_page_text)
                if po_number_match:
                    return po_number_match.group(1)
                
                # Alternative patterns
                patterns = [
                    r'P\.O\.\s*Number:\s*(\d+)',
                    r'Purchase Order Number:\s*(\d+)',
                    r'PO\s*#:\s*(\d+)',
                    r'PO\s*No:\s*(\d+)',
                    r'PO\s*Number:\s*(\d+)'
                ]
                
                for pattern in patterns:
                    match = re.search(pattern, first_page_text, re.IGNORECASE)
                    if match:
                        return match.group(1)
                
                # Fallback: Look for any 7-8 digit number on the right side of the page
                words = first_page_text.split()
                for i, word in enumerate(words):
                    if re.match(r'^\d{7,8}$', word):
                        if i > len(words) / 2:
                            return word
                
                # If still not found, try all pages
                for page in pdf.pages:
                    page_text = page.extract_text() or ""
                    for pattern in patterns:
                        match = re.search(pattern, page_text, re.IGNORECASE)
                        if match:
                            return match.group(1)
                
                # Last resort: look for any 7-8 digit number in the entire document
                full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
                numbers = re.findall(r'\b\d{7,8}\b', full_text)
                if numbers:
                    return numbers[0]
        return ""
    except Exception as e:
        st.error(f"Error extracting PO number: {e}")
        return ""

def extract_so_number_from_wo(pdf_file):
    """Extract SO Number from WO PDF under Product Details section"""
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
        
        # Look for SO Number pattern
        so_number_match = re.search(r"SO Number:\s*([A-Z0-9]+)", full_text)
        if so_number_match:
            return so_number_match.group(1).strip()
        
        # Alternative pattern if the first one fails
        so_number_match_alt = re.search(r"Line Item:\s*\nSO Number:\s*([A-Z0-9]+)", full_text)
        if so_number_match_alt:
            return so_number_match_alt.group(1).strip()
            
        return None
    except Exception as e:
        st.error(f"Error extracting SO Number: {e}")
        return None

def extract_all_so_numbers_from_wo(pdf_file):
    """Extract all SO Numbers from WO PDF (one per WO)"""
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
        
        # Look for SO Number patterns
        so_numbers = []
        
        # Primary pattern: "SO Number:" followed by alphanumeric code
        so_matches = re.findall(r"SO Number:\s*([A-Z0-9]+)", full_text)
        so_numbers.extend(so_matches)
        
        # Alternative pattern: "Line Item:" followed by "SO Number:"
        alt_matches = re.findall(r"Line Item:\s*\nSO Number:\s*([A-Z0-9]+)", full_text)
        so_numbers.extend(alt_matches)
        
        # Another pattern: Look for WO headers and extract SO numbers
        wo_sections = re.split(r'Order Header Details:\s*\*SW\d{8}W\*', full_text)
        
        # Skip the first section (before the first WO header)
        for section in wo_sections[1:]:
            # Look for SO Number pattern in this section
            so_match = re.search(r"SO Number:\s*([A-Z0-9]+)", section)
            if so_match:
                so_numbers.append(so_match.group(1).strip())
            else:
                # Try alternative pattern
                so_match_alt = re.search(r"Line Item:\s*\nSO Number:\s*([A-Z0-9]+)", section)
                if so_match_alt:
                    so_numbers.append(so_match_alt.group(1).strip())
        
        # Remove duplicates while preserving order
        seen = set()
        unique_so_numbers = []
        for so in so_numbers:
            if so not in seen:
                seen.add(so)
                unique_so_numbers.append(so)
        
        return unique_so_numbers
        
    except Exception as e:
        st.error(f"Error extracting SO Numbers: {e}")
        return []  # Always return a list, even on error

def clean_quantity(qty_str):
    """Convert strings like '1,148.0000' or '465.0000' into floats with 4 decimal places"""
    if not qty_str:
        return 0.0000
    qty_str = str(qty_str).strip().replace(",", "")
    try:
        return round(float(qty_str), 4)
    except ValueError:
        return 0.0000

def truncate_after_sri_lanka(addr: str) -> str:
    """
    Extract the address up to and including:
    - The first occurrence of "Sri Lanka" (if found)
    - The second occurrence of "India" (if found twice)
    - The first occurrence of "India" (if found once)
    - Otherwise, the entire address
    Case-insensitive matching.
    """
    if not addr:
        return addr.strip()
    addr_lower = addr.lower()
    sri_lanka_pos = addr_lower.find("sri lanka")
    if sri_lanka_pos != -1:
        return addr[:sri_lanka_pos + len("sri lanka")].strip()
    india_positions = []
    search_pos = 0
    while True:
        india_pos = addr_lower.find("india", search_pos)
        if india_pos == -1:
            break
        india_positions.append(india_pos)
        search_pos = india_pos + 1
    if len(india_positions) >= 2:
        second_india_pos = india_positions[1]
        return addr[:second_india_pos + len("india")].strip()
    if len(india_positions) == 1:
        first_india_pos = india_positions[0]
        return addr[:first_india_pos + len("india")].strip()
    return addr.strip()

def clean_size(size_str):
    """Extract only the primary size from strings like 'S | P' or 'M / M'"""
    if not size_str:
        return ""
    size_str = str(size_str).strip().upper()
    if "|" in size_str:
        return size_str.split("|")[0].strip()
    elif "/" in size_str:
        return size_str.split("/")[0].strip()
    return size_str

def extract_wo_fields(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    delivery = ""
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if "Deliver To:" in ln:
            if i > 0:
                prev_line = lines[i-1].strip()
                if "Customer Delivery Name" in prev_line:
                    delivery = truncate_after_sri_lanka(re.sub(r"Deliver To:\s*", "", ln).strip())
                else:
                    if (any(char.isdigit() for char in prev_line) or 
                        any(indicator in prev_line.lower() for indicator in 
                            ["street", "st", "road", "rd", "avenue", "ave", "building", "block", "no", "#"]) or
                        len(prev_line.split()) > 3):
                        combined = prev_line + " " + re.sub(r"Deliver To:\s*", "", ln).strip()
                        delivery = truncate_after_sri_lanka(combined)
                    else:
                        delivery = truncate_after_sri_lanka(re.sub(r"Deliver To:\s*", "", ln).strip())
            else:
                delivery = truncate_after_sri_lanka(re.sub(r"Deliver To:\s*", "", ln).strip())
            delivery = re.sub(r'PTK ENTERPRISES,\s*', '', delivery, flags=re.IGNORECASE)
            break
    codes = []
    for line in lines:
        if "Product Code" in line:
            found = re.findall(r"Product Code[:\s]*([\w\s\-]+(?:\s*/\s*[\w\s\-]+)*)", line)
            for match in found:
                for code in match.split("/"):
                    clean = code.strip().upper()
                    clean = clean.replace("VSBA", "")
                    clean = clean.replace("-", " ")
                    if clean:
                        codes.append(clean)
            break
    po_numbers = list(set(re.findall(r'\b\d{7,8}\b', text)))
    return {
        "customer_name": "",
        "delivery_address": delivery,
        "product_codes": list(set(codes)),
        "po_numbers": po_numbers
    }

def extract_size_from_cell(cell_text):
    """
    Extract size information from a cell that might contain newlines or special formatting.
    Handles cases like "XS/\nXP", "XL/\nXG", "XXL", etc.
    """
    if not cell_text:
        return ""
    
    # Convert to string and normalize whitespace
    text = str(cell_text).strip()
    
    # If there's a newline, process each line
    if "\n" in text:
        lines = [line.strip() for line in text.split("\n") if line.strip()]
        
        # Check for patterns like ["XS/", "XP"] or ["XL/", "XG"]
        if len(lines) >= 2 and lines[0].endswith("/"):
            # Combine the parts without space
            return lines[0] + lines[1]
        
        # Check for patterns like ["XS", "/XP"] or ["XL", "/XG"]
        if len(lines) >= 2 and lines[1].startswith("/"):
            # Combine the parts without space
            return lines[0] + lines[1]
        
        # If no special pattern, just join with no space
        return "".join(lines)
    
    # If no newline, return as is
    return text

def extract_po_fields(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    lines = [ln.strip() for ln in text.split("\n")]
    capture = False
    address_lines = []
    blank_count = 0
    for ln in lines:
        if "Delivery Location:" in ln:
            capture = True
            continue
        if capture:
            if "Forwarder:" in ln:
                break
            if not ln:
                blank_count += 1
                if blank_count >= 2:
                    break
                continue
            else:
                blank_count = 0
            address_lines.append(ln)
    if address_lines:
        full_address = " ".join(address_lines)
    else:
        plot_found = False
        temp_address_lines = []
        for i, line in enumerate(lines):
            if re.search(r'\b(plot|building)\s*#?\s*\d+', line, re.IGNORECASE):
                plot_found = True
                temp_address_lines = [line]
                for j in range(i + 1, len(lines)):
                    next_line = lines[j].strip()
                    if not next_line:
                        continue
                    temp_address_lines.append(next_line)
                    if re.search(r'\bindia\b', next_line, re.IGNORECASE):
                        break
                    if any(keyword in next_line.lower() for keyword in 
                           ['forwarder', 'shipping', 'invoice', 'payment', 'terms']):
                        temp_address_lines.pop()
                        break
                break
        if plot_found and temp_address_lines:
            full_address = " ".join(temp_address_lines)
        else:
            full_address = ""
    full_address = re.sub(r'\s+', ' ', full_address)
    full_address = full_address.replace('\xa0', ' ')
    full_address = full_address.replace('\u2013', '-')
    full_address = full_address.replace('\u2014', '-')
    final_addr = truncate_after_sri_lanka(full_address)
    final_addr = final_addr.strip()
    po_codes = re.findall(r"(LB\s*\d+)", text)
    sup_ref_codes = re.findall(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z]+[-\s]?\d+)", text, re.IGNORECASE)
    tag_codes = re.findall(r"TAG\.PRC\.TKT_(.*?)_REG", text)
    all_product_codes = list(set([c.strip().upper() for c in sup_ref_codes + tag_codes]))
    return {
        "delivery_location": final_addr,
        "product_codes": po_codes + all_product_codes,
        "all_found_addresses": [final_addr]
    }



def clean_wo_address(addr: str) -> str:
    if not addr:
        return addr.strip()
    addr = re.sub(r',\s*,\s*India\s*$', '', addr, flags=re.IGNORECASE)
    addr = re.sub(r',,\s*', ', ', addr)
    addr = re.sub(r'\s*,\s*', ', ', addr)
    addr = re.sub(r',\s*$', '', addr)
    return addr.strip()

def clean_address_for_comparison(address):
    """
    Clean an address by removing ALL commas (,) and hash symbols (#) for comparison purposes.
    
    Args:
        address (str): The address string to clean
        
    Returns:
        str: The cleaned address with all commas and hash symbols removed
    """
    if not address:
        return ""
    
    # Convert to string first to handle any non-string input
    address = str(address)
    
    # Remove ALL commas and hash symbols
    cleaned = address.replace(",", "").replace("#", "")
    
    # Normalize multiple spaces to single space
    cleaned = " ".join(cleaned.split())
    
    return cleaned.strip()

def compare_addresses(wo, po):
    # Clean BOTH WO and PO addresses before comparison
    wo_name_clean = clean_address_for_comparison(wo["customer_name"])
    wo_addr_clean = clean_address_for_comparison(wo["delivery_address"])
    po_addr_clean = clean_address_for_comparison(po["delivery_location"])
    
    # Compare the CLEANED addresses
    ns = fuzz.token_sort_ratio(wo_name_clean, po_addr_clean)
    as_ = fuzz.token_sort_ratio(wo_addr_clean, po_addr_clean)
    comb = max(ns, as_)
    
    return {
        "WO Name": wo["customer_name"], 
        "WO Addr": wo["delivery_address"], 
        "PO Addr": po["delivery_location"],
        "Name %": ns, 
        "Addr %": as_, 
        "Overall %": comb, 
        "Status": "✅ Match" if comb >= 90 else "⚠️ Review"
    }

def extract_style_numbers(po_pdf_path):
    style_numbers = set()
    with pdfplumber.open(po_pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        if row:
                            for cell in row:
                                if cell:
                                    cell_clean = str(cell).strip().replace(',', '')
                                    if "." in cell_clean:
                                        cell_clean = cell_clean.split(".")[0]
                                    if cell_clean.isdigit():
                                        val = int(cell_clean)
                                        if 10 < val < 100000:
                                            if val > qty:
                                                qty = val
        full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        found_free_text_styles = re.findall(r'\b[A-Z]{2,}\s*\d{3,}\b', full_text)
        style_numbers.update([s.strip() for s in found_free_text_styles])
        if len(pdf.pages) > 0:
            first_page_text = pdf.pages[0].extract_text() or ""
            extracted_section_match = re.search(r'Extracted Style Numbers:\s*(.+)', first_page_text)
            if extracted_section_match:
                extracted_styles = re.findall(r'\b[A-Z]{2,}\s*\d{3,}\b', extracted_section_match.group(1))
                style_numbers.update([s.strip() for s in extracted_styles])
    return list(style_numbers)

def reorder_wo_by_size(wo_items):
    size_order = {"XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5}
    def get_order(wo):
        size = wo.get("Size 1", "").strip().upper()
        return size_order.get(size, 99)
    return sorted(wo_items, key=get_order)

def reorder_po_by_size(po_details):
    size_order = {"XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5}
    def get_order(po):
        size = po.get("Size", "").strip().upper()
        return size_order.get(size, 99)
    return sorted(po_details, key=get_order)

def sort_items_by_size(items):
    size_order = {"XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5}
    def get_size_key(item):
        size = item.get("WO Size", item.get("PO Size", "")).strip().upper()
        return size_order.get(size, 99)
    return sorted(items, key=get_size_key)

def extract_po_details(pdf_file):
    """Enhanced function to handle multiple PO formats with quantity aggregation"""
    pdf_file.seek(0)
    extracted_styles = extract_style_numbers_from_po_first_page(pdf_file)
    repeated_style = extracted_styles[0] if extracted_styles else ""
    pdf_file.seek(0)
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    has_tag_format = "TAG.PRC.TKT_" in text and "Color/Size/Destination :" in text
    has_original_format = any("Colour/Size/Destination:" in line for line in lines) or re.search(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z]+[-\s]?\d+)", text, re.IGNORECASE)
    po_items = []
    item_dict = {}  # Dictionary to aggregate quantities by size, color, and style
    
    # NEW: Extract PO product codes from Item column using TAG.HANG pattern
    po_product_codes_from_item = extract_po_product_codes_from_tag_hang_pattern(pdf_file)
    
    if has_tag_format and not has_original_format:
        tag_match = re.search(r"TAG\.PRC\.TKT_(.*?)_REG", text)
        product_code_used = tag_match.group(1).strip().upper() if tag_match else ""
        product_code_used = product_code_used.replace("-", " ")
        i = 0
        while i < len(lines):
            line = lines[i]
            item_match = re.match(r'^(\d+)\s+TAG\.PRC\.TKT_.*?([\d,]+\.\d+)\s+PCS', line)
            if item_match:
                item_no = item_match.group(1)
                quantity_str = item_match.group(2)
                quantity = clean_quantity(quantity_str)
                colour = size = ""
                for j in range(i + 1, min(i + 5, len(lines))):
                    next_line = lines[j]
                    if "Color/Size/Destination :" in next_line:
                        cs_part = next_line.split(":", 1)[1].strip()
                        cs_parts = [part.strip() for part in cs_part.split(" / ") if part.strip()]
                        if len(cs_parts) >= 2:
                            size_keywords = ["XS", "S", "M", "L", "XL", "XXL", "XXXL", "XXG", "P", "G"]
                            first_part = cs_parts[0].strip()
                            if first_part.upper() in size_keywords:
                                size = first_part.upper()
                                colour = cs_parts[1].split()[0] if cs_parts[1] else ""
                            else:
                                colour_part = cs_parts[0].strip()
                                colour = colour_part.split()[0] if colour_part else ""
                                size = cs_parts[1].strip().upper()
                        
                        # NEW LOGIC: Check for size before last slash
                        # Find the last occurrence of a slash
                        last_slash_index = cs_part.rfind('/')
                        if last_slash_index != -1:
                            # Get the substring before the last slash
                            before_slash = cs_part[:last_slash_index].strip()
                            # Split into tokens and check in reverse order
                            tokens = before_slash.split()
                            for token in reversed(tokens):
                                if token.upper() in size_keywords:
                                    size = token.upper()
                                    break
                        break
                item_key = (size, colour.upper() if colour else "", repeated_style)
                if item_key in item_dict:
                    item_dict[item_key]["Quantity"] += quantity
                else:
                    item_dict[item_key] = {
                        "Item_Number": item_no,
                        "Item_Code": f"TAG_{product_code_used}",
                        "Quantity": quantity,
                        "Colour_Code": colour.upper() if colour else "",
                        "Size": size,
                        "Style 2": repeated_style,
                        "Product_Code": product_code_used,
                    }
            i += 1
    else:
        # ORIGINAL FORMAT HANDLING (kept original logic)
        sup_ref_match = re.search(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z]+[-\s]?\d+)", text, re.IGNORECASE)
        sup_ref_code = sup_ref_match.group(1).strip().upper() if sup_ref_match else ""
        sup_ref_code = sup_ref_code.replace("-", " ")
        tag_code = ""
        for i, line in enumerate(lines):
            if "Item Description" in line:
                if i + 2 < len(lines):
                    second_line = lines[i + 2]
                    match = re.search(r"TAG\.PRC\.TKT_(.*?)_REG", second_line)
                    if match:
                        tag_code = match.group(1).strip().upper()
                        tag_code = tag_code.replace("-", " ")
                break
        product_code_used = sup_ref_code if sup_ref_code else tag_code
        for i, line in enumerate(lines):
            # Strict pattern (original)
            item_match = re.match(r'^(\d+)\s+([A-Z0-9]+)\s+(\d+)\s+([\d,]+\.\d+)\s+PCS', line)
            # Fallback pattern (only used if strict fails) — non-breaking
            relaxed_match = None
            if not item_match:
                relaxed_match = re.match(r'^(\d+)\s+([A-Z0-9]+)\b.*?([\d,]+(?:\.\d+)?)\s+PCS', line)
            if item_match or relaxed_match:
                if item_match:
                    item_no, item_code, _, qty_str = item_match.groups()
                else:
                    item_no, item_code, qty_str = relaxed_match.groups()
                quantity = clean_quantity(qty_str)
                colour = size = ""
                # Keep original logic, just allow "Color/Size/Destination :" as additional fall-back
                for j in range(i + 1, min(i + 10, len(lines))):
                    ln = lines[j]
                    if not colour and ("Colour/Size/Destination:" in ln or "Color/Size/Destination :" in ln):
                        cs = ln.split(":", 1)[1].strip()
                        size_keywords = ["XS", "S", "M", "L", "XL", "XXL", "XXXL", "XXG", "P", "G"]
                        parts = [p.strip() for p in cs.split("/") if p.strip()]
                        if parts:
                            size_part = parts[0].split("|")[0].strip().upper()
                            if size_part in size_keywords:
                                size = size_part
                                if len(parts) > 1:
                                    colour = parts[1].strip().split()[0].strip().upper()
                            else:
                                colour = re.match(r'^(\S+)', cs).group(1) if re.match(r'^(\S+)', cs) else ""
                                size_match = re.search(r'/\s*([^/]+)\s*/', cs)
                                if size_match:
                                    size = size_match.group(1).strip()
                        
                        # NEW LOGIC: Check for size before last slash
                        # Find the last occurrence of a slash
                        last_slash_index = cs.rfind('/')
                        if last_slash_index != -1:
                            # Get the substring before the last slash
                            before_slash = cs[:last_slash_index].strip()
                            # Split into tokens and check in reverse order
                            tokens = before_slash.split()
                            for token in reversed(tokens):
                                if token.upper() in size_keywords:
                                    size = token.upper()
                                    break
                item_key = (size.upper() if size else "", (colour or "").strip().upper(), repeated_style)
                if item_key in item_dict:
                    item_dict[item_key]["Quantity"] += quantity
                else:
                    item_dict[item_key] = {
                        "Item_Number": item_no,
                        "Item_Code": item_code,
                        "Quantity": quantity,
                        "Colour_Code": (colour or "").strip().upper(),
                        "Size": (size or "").strip().upper(),
                        "Style 2": repeated_style,
                        "Product_Code": product_code_used,
                    }
    po_items = list(item_dict.values())
    
    # NEW: Add PO product codes from Item column to the returned data
    # This will be used in the Product Code Analysis section
    return {
        "po_items": po_items,
        "po_product_codes_from_item": po_product_codes_from_item
    }
def extract_wo_items_table(pdf_file, product_codes=None):
    """
    Enhanced function to extract WO items from Victoria's Secret price ticket tables
    with improved table detection and data extraction for all formats, including sizes split across lines
    """
    items = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # First try standard table extraction
            tables = page.extract_tables()
            
            # If standard extraction fails, try with explicit lines
            if not tables or len(tables) == 0:
                tables = page.extract_tables({
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "explicit_vertical_lines": page.curves + page.edges,
                    "explicit_horizontal_lines": page.curves + page.edges,
                })
            
            # Process each table
            for table_idx, table in enumerate(tables):
                if not table or len(table) < 2:
                    continue
                
                # Pre-process table to handle sizes split across cells and within cells
                processed_table = []
                for row in table:
                    if not row:
                        continue
                    
                    processed_row = []
                    i = 0
                    while i < len(row):
                        cell = str(row[i]) if row[i] is not None else ""
                        
                        # Check if this cell ends with a slash and the next cell contains a size suffix
                        if i < len(row) - 1 and cell.strip().endswith("/"):
                            next_cell = str(row[i+1]) if row[i+1] is not None else ""
                            # Check if next cell is a size suffix (XP, P, M, G, XG)
                            if next_cell.strip().upper() in ["XP", "P", "M", "G", "XG"]:
                                # Combine the cells
                                combined_cell = cell + next_cell
                                processed_row.append(combined_cell)
                                i += 2  # Skip the next cell
                                continue
                        
                        # If not a split size, just add the cell as-is
                        processed_row.append(cell)
                        i += 1
                    
                    processed_table.append(processed_row)
                
                # Now find the header row
                header_row_idx = -1
                column_positions = {}
                
                for i, row in enumerate(processed_table):
                    if not row:
                        continue
                    
                    row_text = " ".join([str(cell).strip() for cell in row if cell])
                    if any(term in row_text for term in ["Style", "Colour Code", "Size", "Quantity"]):
                        header_row_idx = i
                        
                        # Map column positions
                        for j, cell in enumerate(row):
                            cell_text = str(cell).strip().lower() if cell else ""
                            if "style" in cell_text:
                                column_positions["style"] = j
                            elif "colour" in cell_text or "color" in cell_text:
                                column_positions["color_code"] = j
                            elif "size 1" in cell_text or "size" in cell_text:
                                column_positions["size1"] = j
                            elif "size 2" in cell_text:
                                column_positions["size2"] = j
                            elif "panty" in cell_text:
                                column_positions["panty_length"] = j
                            elif "retail" in cell_text and "us" in cell_text:
                                column_positions["retail_us"] = j
                            elif "retail" in cell_text and "ca" in cell_text:
                                column_positions["retail_ca"] = j
                            elif "multi" in cell_text:
                                column_positions["multi_price"] = j
                            elif "sku" in cell_text:
                                column_positions["sku"] = j
                            elif "article" in cell_text:
                                column_positions["article"] = j
                            elif "quantity" in cell_text or "qty" in cell_text:
                                column_positions["quantity"] = j
                        break
                
                # If we couldn't find a header row, try to infer it
                if header_row_idx == -1:
                    for i, row in enumerate(processed_table):
                        if not row or len(row) < 8:
                            continue
                        
                        first_cell = str(row[0]).strip()
                        if re.match(r'^\d{8}$', first_cell):
                            has_size = False
                            for cell in row:
                                cell_str = str(cell).strip().upper()
                                # Check for any size format, including combined ones
                                if any(size in cell_str for size in ["XS/XP", "S/P", "M/M", "L/G", "XL/XG", "XXL", "XXXL", "XXG", "XG", "XS", "S", "M", "L", "XL", "P", "G"]):
                                    has_size = True
                                    break
                            
                            if has_size:
                                header_row_idx = i
                                column_positions = {
                                    "style": 0,
                                    "color_code": 1,
                                    "size1": 2,
                                    "size2": 3,
                                    "panty_length": 4,
                                    "retail_us": 5,
                                    "retail_ca": 6,
                                    "multi_price": 7,
                                    "sku": 8,
                                    "article": 9,
                                    "quantity": len(row) - 1
                                }
                                break
                
                # Skip if we couldn't determine the header
                if header_row_idx == -1:
                    continue
                
                # Process data rows
                for row_idx, row in enumerate(processed_table[header_row_idx + 1:], header_row_idx + 1):
                    if not row or len(row) < max(column_positions.values()) + 1:
                        continue
                    
                    try:
                        style = str(row[column_positions.get("style", 0)] or "").strip()
                        color_code = str(row[column_positions.get("color_code", 1)] or "").strip().upper()
                        
                        # Extract size1 with special handling for multi-line cells
                        size1_raw = str(row[column_positions.get("size1", 2)] or "")
                        size1 = extract_size_from_cell(size1_raw)

                        # If we didn't get a valid size, try to find it in other cells
                        if not size1 or not any(size in size1.upper() for size in ["XS", "S", "M", "L", "XL", "XXL", "XXXL", "XXG", "XG", "P", "G"]):
                            # Check each cell for size patterns
                            for cell in row:
                                cell_str = str(cell) if cell is not None else ""
                                extracted_size = extract_size_from_cell(cell_str)
                                if any(size in extracted_size.upper() for size in ["XS", "S", "M", "L", "XL", "XXL", "XXXL", "XXG", "XG", "P", "G"]):
                                    size1 = extracted_size
                                    break
                        
                        # Check if the size cell contains a newline (like "XS\nXP")
                        if "\n" in size1_raw:
                            # Split by newline and take the first part
                            size_parts = size1_raw.split("\n")
                            # Process each part to handle the case where one part is just "/" and the next is "XP"
                            processed_size = ""
                            for part in size_parts:
                                if part.strip() == "/":
                                    processed_size += "/"
                                else:
                                    processed_size += part.strip()
                            
                            # Now clean the processed size
                            size1 = clean_size(processed_size)
                        else:
                            size1 = clean_size(size1_raw)
                        
                        # If we didn't get a valid size, try to find it in other cells
                        if not size1:
                            # Check each cell for size patterns
                            for cell in row:
                                cell_str = str(cell).strip()
                                
                                # Check if the cell contains a newline
                                if "\n" in cell_str:
                                    # Split by newline and check each part
                                    parts = cell_str.split("\n")
                                    processed_cell = ""
                                    for part in parts:
                                        if part.strip() == "/":
                                            processed_cell += "/"
                                        else:
                                            processed_cell += part.strip()
                                    
                                    # Look for size patterns in the processed cell
                                    cell_upper = processed_cell.upper()
                                    size_match = re.search(r'\b(XS/XP|S/P|M/M|L/G|XL/XG|XXL|XXXL|XXG|XG|XS|S|M|L|XL|P|G)\b', cell_upper)
                                    if size_match:
                                        size1 = clean_size(size_match.group(1))
                                        break
                                else:
                                    # Look for size patterns in the whole cell
                                    cell_upper = cell_str.upper()
                                    size_match = re.search(r'\b(XS/XP|S/P|M/M|L/G|XL/XG|XXL|XXXL|XXG|XG|XS|S|M|L|XL|P|G)\b', cell_upper)
                                    if size_match:
                                        size1 = clean_size(size_match.group(1))
                                        break
                        
                        # Extract quantity
                        quantity_str = ""
                        if "quantity" in column_positions:
                            quantity_str = str(row[column_positions["quantity"]] or "").strip()
                        
                        if not quantity_str or not re.search(r'\d', quantity_str):
                            # Try the last column as a fallback
                            quantity_str = str(row[-1] or "").strip()
                        
                        quantity = clean_quantity(quantity_str)
                        
                        # Only add item if we have valid style and quantity
                        if style and quantity > 0:
                            item_data = {
                                "Style": style,
                                "WO Colour Code": color_code,
                                "Size 1": size1,
                                "Quantity": quantity,
                                "WO Product Code": " / ".join(product_codes) if product_codes else ""
                            }
                            
                            # Extract size2 with special handling for multi-line cells
                            if "size2" in column_positions:
                                size2_raw = str(row[column_positions["size2"]] or "")
                                item_data["Size 2"] = extract_size_from_cell(size2_raw)
                                if "\n" in size2_raw:
                                    # Split by newline and process each part
                                    size_parts = size2_raw.split("\n")
                                    processed_size = ""
                                    for part in size_parts:
                                        if part.strip() == "/":
                                            processed_size += "/"
                                        else:
                                            processed_size += part.strip()
                                    
                                    item_data["Size 2"] = clean_size(processed_size)
                                else:
                                    item_data["Size 2"] = clean_size(size2_raw)
                            
                            if "panty_length" in column_positions:
                                item_data["Panty Length"] = str(row[column_positions["panty_length"]] or "").strip()
                            
                            if "retail_us" in column_positions:
                                item_data["Retail US"] = str(row[column_positions["retail_us"]] or "").strip()
                            
                            if "retail_ca" in column_positions:
                                item_data["Retail CA"] = str(row[column_positions["retail_ca"]] or "").strip()
                            
                            if "multi_price" in column_positions:
                                item_data["Multi Price"] = str(row[column_positions["multi_price"]] or "").strip()
                            
                            if "sku" in column_positions:
                                item_data["SKU"] = str(row[column_positions["sku"]] or "").strip()
                            
                            if "article" in column_positions:
                                item_data["Article"] = str(row[column_positions["article"]] or "").strip()
                            
                            items.append(item_data)
                    
                    except (ValueError, IndexError):
                        continue
    
    # If we still don't have items, try text-based extraction
    if not items:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
        
        # Try the existing pattern first
        lines = full_text.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Existing pattern - keep this unchanged
            normalized_line = re.sub(r'\s+', ' ', line)
            pattern = r'(\d{8})\s+([A-Z0-9]+)\s+([A-Z]{2}(?:/[A-Z]{2})?)\s+\$[\d.]+\s+\$[\d.]+\s+\d+\s+\d+\s+(\d{1,4}(?:,\d{3})*)'
            match = re.search(pattern, normalized_line)

            if match:
                style = match.group(1)
                color_code = match.group(2)
                size = match.group(3)
                quantity = match.group(4).replace(',', '')
                
                try:
                    items.append({
                        "Style": style,
                        "WO Colour Code": color_code.upper(),
                        "Size 1": size,
                        "Quantity": int(quantity),
                        "WO Product Code": " / ".join(product_codes) if product_codes else ""
                    })
                except ValueError:
                    continue
        
        # If still no items, try the new pattern for this specific WO format
        if not items:
            # Split the text into work order sections
            wo_sections = re.split(r'Order Header Details: \*SW(\d{8})W\*', full_text)
            
            # Skip the first section (before the first WO header)
            for section in wo_sections[1:]:
                # Find the table section
                table_start = section.find('Size/Age Breakdown:')
                if table_start == -1:
                    continue
                
                table_text = section[table_start + len('Size/Age Breakdown:'):].strip()
                
                # Split into lines and process
                table_lines = [line.strip() for line in table_text.split('\n') if line.strip()]
                
                i = 0
                while i < len(table_lines):
                    line = table_lines[i]
                    
                    # Check if this line starts with a style number (8 digits)
                    style_match = re.match(r'^(\d{8})\s+([A-Z0-9]+)\s*(.*)', line)
                    if style_match:
                        style = style_match.group(1)
                        color_code = style_match.group(2)
                        rest = style_match.group(3)
                        
                        # Handle size that might be split across lines
                        size = ""
                        quantity = ""
                        
                        # Try to find size in the current line
                        size_match = re.search(r'([A-Z]{2}(?:/[A-Z]{2})?)', rest)
                        if size_match:
                            size = size_match.group(1)
                            # Look for quantity in the rest of the line
                            qty_match = re.search(r'(\d{1,4}(?:,\d{3})*)\s*$', rest[size_match.end():])
                            if qty_match:
                                quantity = qty_match.group(1)
                        
                        # If we didn't find both size and quantity, check the next line
                        if (not size or not quantity) and i+1 < len(table_lines):
                            next_line = table_lines[i+1]
                            
                            # If size wasn't found, look for it in the next line
                            if not size:
                                next_size_match = re.search(r'^([A-Z]{2}(?:/[A-Z]{2})?)', next_line)
                                if next_size_match:
                                    size = next_size_match.group(1)
                                    # Look for quantity in the rest of the next line
                                    next_qty_match = re.search(r'(\d{1,4}(?:,\d{3})*)\s*$', next_line[next_size_match.end():])
                                    if next_qty_match:
                                        quantity = next_qty_match.group(1)
                            # If size was found but quantity wasn't, look for quantity in the next line
                            elif not quantity:
                                next_qty_match = re.search(r'(\d{1,4}(?:,\d{3})*)\s*$', next_line)
                                if next_qty_match:
                                    quantity = next_qty_match.group(1)
                            
                            # If we found something in the next line, skip it
                            if (size and quantity) or (not size and next_size_match) or (not quantity and next_qty_match):
                                i += 1
                        
                        # If we found both size and quantity, add the item
                        if size and quantity:
                            try:
                                items.append({
                                    "Style": style,
                                    "WO Colour Code": color_code.upper(),
                                    "Size 1": size,
                                    "Quantity": int(quantity.replace(',', '')),
                                    "WO Product Code": " / ".join(product_codes) if product_codes else ""
                                })
                            except ValueError:
                                pass
                    
                    i += 1
                
    # Aggregate quantities for items with same style, color code, and size
    aggregated_items = {}
    for item in items:
        key = (item["Style"], item["WO Colour Code"], item["Size 1"])
        if key in aggregated_items:
            aggregated_items[key]["Quantity"] += item["Quantity"]
        else:
            aggregated_items[key] = item.copy()
    
    # Convert back to list
    final_items = list(aggregated_items.values())
    
    return final_items


def _extract_qty_from_text(text, min_value=1, max_value=100000):
    if not text:
        return None
    candidates = re.findall(r'\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?', str(text))
    best = None
    for token in candidates:
        num_str = token.replace(",", "")
        try:
            int_part = num_str.split(".", 1)[0]
            if not int_part:
                continue
            val = int(int_part)
            if min_value <= val <= max_value:
                if best is None or val > best:
                    best = val
        except ValueError:
            continue
    return best

def _dedupe(items):
    seen = set()
    unique = []
    for it in items:
        key = (it["Style"], it.get("WO Colour Code", ""), it.get("Size 1", ""), it["Quantity"])
        if key in seen:
            continue
        seen.add(key)
        unique.append(it)
    return unique


def compare_excel_style_with_po_style2(wo_pdf_file, po_pdf_file, excel_file=None):
    """
    Compare style numbers from WO with Style 2 column from PO.
    Also uses style numbers from uploaded Excel sheet as reference.
    
    Args:
        wo_pdf_file: WO PDF file object
        po_pdf_file: PO PDF file object
        excel_file: Optional Excel file with style numbers
        
    Returns:
        DataFrame with style comparison results
    """
    # Extract items from WO
    wo_items = extract_wo_items_table(wo_pdf_file)
    
    # Extract items from PO
    po_details = extract_po_details(po_pdf_file)
    po_items = po_details.get("po_items", [])
    
    # Get style numbers from WO Style column (all styles, not just numeric)
    wo_styles = []
    for item in wo_items:
        style = item.get("Style", "")
        if style:  # Get all styles, not just numeric ones
            wo_styles.append(style)
    
    # Get style numbers from PO Style 2 column
    po_styles = []
    for item in po_items:
        style2 = item.get("Style 2", "")
        if style2:
            po_styles.append(style2)
    
    # Get style numbers from Excel sheet if provided
    excel_styles = []
    if excel_file:
        try:
            # Read the Excel file
            excel_df = pd.read_excel(excel_file)
            
            # Look for a column that might contain style numbers
            style_col = None
            for col in excel_df.columns:
                if "style" in col.lower():
                    style_col = col
                    break
            
            if style_col:
                # Extract unique style numbers from this column
                for style in excel_df[style_col].dropna().unique():
                    if str(style).strip():  # Make sure it's not empty
                        excel_styles.append(str(style).strip())
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
    
    # Remove duplicates
    wo_styles = list(set(wo_styles))
    po_styles = list(set(po_styles))
    excel_styles = list(set(excel_styles))
    
    # Create comparison data
    comparison_data = []
    
    # Create a mapping of all style numbers for easy lookup
    all_styles = set(wo_styles + po_styles + excel_styles)
    
    # Find matches between WO and PO
    for style in wo_styles:
        if style in po_styles:
            # Style found in both WO and PO
            comparison_data.append({
                "Style Number": style,
                "WO Style Column": "✅ Found",
                "PO Style 2 Column": "✅ Found",
                "Excel Style": "✅ Found" if style in excel_styles else "❌ Not Found",
                "Match Status": "✅ Perfect Match"
            })
        else:
            # Style only in WO
            comparison_data.append({
                "Style Number": style,
                "WO Style Column": "✅ Found",
                "PO Style 2 Column": "❌ Not Found",
                "Excel Style": "✅ Found" if style in excel_styles else "❌ Not Found",
                "Match Status": "⚠️ Only in WO"
            })
    
    # Find matches between PO and Excel
    for style in po_styles:
        if style in excel_styles:
            # Style found in both PO and Excel
            comparison_data.append({
                "Style Number": style,
                "WO Style Column": "❌ Not Found",
                "PO Style 2 Column": "✅ Found",
                "Excel Style": "✅ Found",
                "Match Status": "✅ Perfect Match"
            })
        else:
            # Style only in PO
            comparison_data.append({
                "Style Number": style,
                "WO Style Column": "❌ Not Found",
                "PO Style 2 Column": "✅ Found",
                "Excel Style": "✅ Found",
                "Match Status": "⚠️ Only in PO"
            })
    
    # Find matches between Excel and WO
    for style in excel_styles:
        if style in wo_styles:
            # Style found in both Excel and WO
            comparison_data.append({
                "Style Number": style,
                "WO Style Column": "✅ Found",
                "PO Style 2 Column": "❌ Not Found",
                "Excel Style": "✅ Found",
                "Match Status": "✅ Perfect Match"
            })
        else:
            # Style only in Excel
            comparison_data.append({
                "Style Number": style,
                "WO Style Column": "❌ Not Found",
                "PO Style 2 Column": "❌ Not Found",
                "Excel Style": "✅ Found",
                "Match Status": "⚠️ Only in Excel"
            })
    
    # Create DataFrame
    df = pd.DataFrame(comparison_data)
    
    # Sort by match status and style number
    if not df.empty:
        df = df.sort_values(by=["Match Status", "Style Number"])
    
    return df

def display_excel_style_comparison(wo_pdf_file, po_pdf_file, excel_file=None):
    """
    Display the style comparison between WO, PO, and Excel in Streamlit.
    
    Args:
        wo_pdf_file: WO PDF file object
        po_pdf_file: PO PDF file object
        excel_file: Optional Excel file with style numbers
    """
    st.subheader("Style Number Comparison: WO vs PO vs Excel")
    
    # Get comparison data
    comparison_df = compare_excel_style_with_po_style2(wo_pdf_file, po_pdf_file, excel_file)
    
    if comparison_df.empty:
        st.info("No style numbers found for comparison.")
        return
    
    # Display summary statistics
    total_styles = len(comparison_df)
    matched_styles = len(comparison_df[comparison_df["Match Status"] == "✅ Perfect Match"])
    wo_only_styles = len(comparison_df[comparison_df["Match Status"] == "⚠️ Only in WO"])
    po_only_styles = len(comparison_df[comparison_df["Match Status"] == "⚠️ Only in PO"])
    excel_only_styles = len(comparison_df[comparison_df["Match Status"] == "⚠️ Only in Excel"])
    
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Total Styles", total_styles)
    with col2:
        st.metric("Perfect Match", matched_styles)
    with col3:
        st.metric("Only in WO", wo_only_styles)
    with col4:
        st.metric("Only in PO", po_only_styles)
    with col5:
        st.metric("Only in Excel", excel_only_styles)
    
    # Display the comparison table
    st.dataframe(comparison_df, use_container_width=True)
    
    # Color-code the status for better visibility
    def highlight_status(val):
        if "✅" in val:
            return 'background-color: #d4edda'
        elif "⚠️" in val:
            return 'background-color: #fff3cd'
        elif "❌" in val:
            return 'background-color: #f8d7da'
        return ''
    
    # Display styled table
    styled_df = comparison_df.style.applymap(highlight_status, subset=['Match Status'])
    st.dataframe(styled_df, use_container_width=True)

def get_excel_style_numbers(excel_file):
    """
    Extract style numbers from an Excel file.
    
    Args:
        excel_file: Excel file object
        
    Returns:
        List of unique style numbers
    """
    if not excel_file:
        return []
    
    try:
        # Read the Excel file
        excel_df = pd.read_excel(excel_file)
        
        # Look for a column that might contain style numbers
        style_col = None
        for col in excel_df.columns:
            if "style" in col.lower():
                style_col = col
                break
        
        if style_col:
            # Extract unique style numbers from this column
            styles = []
            for style in excel_df[style_col].dropna().unique():
                if str(style).strip():  # Make sure it's not empty
                    styles.append(str(style).strip())
            return styles
        else:
            # If no style column found, try common column names
            common_style_columns = ["Style", "Style Number", "Style No", "Style Code"]
            for col in common_style_columns:
                if col in excel_df.columns:
                    styles = []
                    for style in excel_df[col].dropna().unique():
                        if str(style).strip():  # Make sure it's not empty
                            styles.append(str(style).strip())
                    if styles:
                        return styles
            return []
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return []

def extract_wo_items_table_enhanced(pdf_file, product_codes=None):
    return extract_wo_items_table(pdf_file, product_codes)
def debug_po_extraction(pdf_file):
    """Debug function to extract and display PO address information"""
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

def debug_extract_product_codes(pdf_file):
    """
    Debug function to show the extraction process for product codes
    """
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                st.write(f"### Page {page_num+1}")
                
                # Extract tables from the page
                tables = page.extract_tables()
                
                if not tables:
                    st.write("No tables found on this page")
                    continue
                
                # Process each table
                for table_idx, table in enumerate(tables):
                    st.write(f"#### Table {table_idx+1}")
                    
                    # Display the table
                    df = pd.DataFrame(table)
                    st.dataframe(df, use_container_width=True)
                    
                    # Find the header row to locate the "Item" column
                    header_row_idx = -1
                    item_col_idx = -1
                    
                    for i, row in enumerate(table):
                        if not row:
                            continue
                        
                        # Check if this row contains "Item" in the first cell
                        first_cell = str(row[0]).strip() if row[0] else ""
                        if "Item" in first_cell:
                            header_row_idx = i
                            st.write(f"Found header row at index {i}")
                            
                            # Find the index of the "Item" column
                            for j, cell in enumerate(row):
                                cell_text = str(cell).strip().lower() if cell else ""
                                if "item" in cell_text:
                                    item_col_idx = j
                                    st.write(f"Found Item column at index {j}")
                                    break
                            break
                    
                    # If we found the header and Item column, process the data rows
                    if header_row_idx != -1 and item_col_idx != -1:
                        st.write("Processing data rows:")
                        for row_idx, row in enumerate(table[header_row_idx + 1:], header_row_idx + 1):
                            if not row or len(row) <= item_col_idx:
                                continue
                            
                            # Get the Item cell value
                            item_cell = str(row[item_col_idx]).strip() if row[item_col_idx] else ""
                            st.write(f"Row {row_idx}: Item cell = '{item_cell}'")
                            
                            # Check if the item cell matches the pattern TAG.HANG_[PRODUCT_CODE]_TAGPRCTKT_
                            pattern_match = re.match(r'TAG\.HANG_(.+?)_TAGPRCTKT_', item_cell)
                            if pattern_match:
                                product_code = pattern_match.group(1).strip()
                                st.write(f"  ✓ Extracted product code: {product_code}")
                            else:
                                st.write(f"  ✗ No match found")
    
    except Exception as e:
        st.error(f"Error in debug extraction: {e}")

def extract_product_codes_from_item_column(pdf_file):
    """
    Extract product codes from the Item column in PO PDFs with format TAG.HANG_[PRODUCT_CODE]_TAGPRCTKT_
    
    Args:
        pdf_file: PDF file object
        
    Returns:
        List of product codes extracted from between TAG.HANG_ and _TAGPRCTKT_
    """
    product_codes = []
    
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                # Extract tables from the page
                tables = page.extract_tables()
                
                # Process each table
                for table in tables:
                    if not table:
                        continue
                    
                    # Find the header row to locate the "Item" column
                    header_row_idx = -1
                    item_col_idx = -1
                    
                    for i, row in enumerate(table):
                        if not row:
                            continue
                        
                        # Check if this row contains "Item" in the first cell
                        first_cell = str(row[0]).strip() if row[0] else ""
                        if "Item" in first_cell:
                            header_row_idx = i
                            
                            # Find the index of the "Item" column
                            for j, cell in enumerate(row):
                                cell_text = str(cell).strip().lower() if cell else ""
                                if "item" in cell_text:
                                    item_col_idx = j
                                    break
                            break
                    
                    # If we found the header and Item column, process the data rows
                    if header_row_idx != -1 and item_col_idx != -1:
                        for row in table[header_row_idx + 1:]:
                            if not row or len(row) <= item_col_idx:
                                continue
                            
                            # Get the Item cell value
                            item_cell = str(row[item_col_idx]).strip() if row[item_col_idx] else ""
                            
                            # Check if the item cell matches the pattern TAG.HANG_[PRODUCT_CODE]_TAGPRCTKT_
                            pattern_match = re.match(r'TAG\.HANG_(.+?)_TAGPRCTKT_', item_cell)
                            if pattern_match:
                                product_code = pattern_match.group(1).strip()
                                if product_code:
                                    product_codes.append(product_code)
    
    except Exception as e:
        st.error(f"Error extracting product codes: {e}")
    
    # Remove duplicates while preserving order
    seen = set()
    unique_product_codes = []
    for code in product_codes:
        if code not in seen:
            seen.add(code)
            unique_product_codes.append(code)
    
    return unique_product_codes

def debug_item_column_extraction(pdf_file):
    """Debug function to see what's in the Item column"""
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                print(f"### Page {page_num + 1} Analysis")  # Using print instead of st.write
                
                # Extract tables from the page
                tables = page.extract_tables()
                
                if not tables:
                    print("No tables found on this page")
                    continue
                
                for table_num, table in enumerate(tables):
                    print(f"#### Table {table_num + 1}")
                    
                    # Find the header row to locate the "Item" column
                    header_row_idx = -1
                    item_col_idx = -1
                    
                    for i, row in enumerate(table):
                        if not row:
                            continue
                        
                        # Check if this row contains "Item" in the first cell
                        first_cell = str(row[0]).strip() if row[0] else ""
                        if "Item" in first_cell:
                            header_row_idx = i
                            print(f"Found header at row {i}")
                            
                            # Find the index of the "Item" column
                            for j, cell in enumerate(row):
                                cell_text = str(cell).strip().lower() if cell else ""
                                if "item" in cell_text:
                                    item_col_idx = j
                                    print(f"Item column found at index {j}")
                                    print(f"Header row: {row}")
                                    break
                            break
                    
                    # If we found the header and Item column, process the data rows
                    if header_row_idx != -1 and item_col_idx != -1:
                        print("Processing data rows:")
                        for row_num, row in enumerate(table[header_row_idx + 1:], header_row_idx + 1):
                            if not row or len(row) <= item_col_idx:
                                continue
                            
                            # Get the Item cell value
                            item_cell = str(row[item_col_idx]).strip() if row[item_col_idx] else ""
                            print(f"Row {row_num}: Item cell = '{item_cell}'")
                            
                            # Check if the item cell matches the pattern
                            if "TAG.HANG_" in item_cell and "_TAGPRCTKT_" in item_cell:
                                print(f"  ✓ Contains TAG.HANG and _TAGPRCTKT_")
                                pattern_match = re.match(r'TAG\.HANG_(.+?)_TAGPRCTKT_', item_cell)
                                if pattern_match:
                                    product_code = pattern_match.group(1).strip()
                                    print(f"  ✓ Extracted product code: '{product_code}'")
                                else:
                                    print(f"  ✗ Pattern match failed")
                            else:
                                print(f"  ✗ Does not contain required pattern")
                    else:
                        print("Could not find Item column or header")
                    
                    print("---")
    
    except Exception as e:
        print(f"Error in debug: {e}")  # Using print instead of st.error
    
    return None  # Return None since we're just debugging

def extract_po_product_codes_from_tag_hang_pattern(pdf_file):
    """
    Extract product codes from PO PDF using TAG.HANG pattern.
    This function looks for patterns like "TAG.HANG_ABC123_TAGPRCTKT" in the text.
    """
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        
        # Extract TAG.HANG patterns
        tag_hang_codes = re.findall(r'TAG\.HANG_(.*?)_TAGPRCTKT', text)
        return tag_hang_codes
    except Exception as e:
        st.error(f"Error extracting PO product codes from TAG.HANG pattern: {e}")
        return []

def extract_all_po_product_codes(pdf_file):
    """
    Extract all product codes from a PO PDF file using multiple patterns.
    Returns a list of unique product codes found in the PO.
    """
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        
        all_codes = []
        
        # Pattern 1: LB followed by digits (e.g., LB12345)
        lb_codes = re.findall(r'LB\s*(\d+)', text, re.IGNORECASE)
        all_codes.extend([f"LB{code}" for code in lb_codes])
        
        # Pattern 2: Sup. Ref. patterns (e.g., Sup. Ref. : ABC123)
        sup_ref_codes = re.findall(r'Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z0-9]+)', text, re.IGNORECASE)
        all_codes.extend(sup_ref_codes)
        
        # Pattern 3: TAG.PRC.TKT patterns (e.g., TAG.PRC.TKT_ABC123_REG)
        tag_codes = re.findall(r'TAG\.PRC\.TKT_(.*?)_REG', text)
        all_codes.extend(tag_codes)
        
        # NEW Pattern: TAG.HANG patterns (e.g., TAG.HANG_ABC123_TAGPRCTKT)
        tag_hang_codes = re.findall(r'TAG\.HANG_(.*?)_TAGPRCTKT', text)
        all_codes.extend(tag_hang_codes)
        
        # Pattern 4: Extract from Item column using TAG.HANG pattern
        tag_hang_item_codes = extract_po_product_codes_from_tag_hang_pattern(pdf_file)
        all_codes.extend(tag_hang_item_codes)
        
        # Pattern 5: Any 8-digit numbers (common style numbers)
        style_numbers = re.findall(r'\b\d{8}\b', text)
        all_codes.extend(style_numbers)
        
        # Pattern 6: Alphanumeric codes with specific patterns
        # Look for codes like ABC-123, ABC123, etc.
        pattern_codes = re.findall(r'\b[A-Z]{2,}\d{2,}\b', text)
        all_codes.extend(pattern_codes)
        
        # Clean and deduplicate
        cleaned_codes = []
        seen = set()
        for code in all_codes:
            # Clean the code: remove extra spaces, convert to uppercase
            cleaned = code.strip().upper()
            # Remove any non-alphanumeric characters except dash
            cleaned = re.sub(r'[^A-Z0-9\-]', '', cleaned)
            if cleaned and cleaned not in seen:
                cleaned_codes.append(cleaned)
                seen.add(cleaned)
        
        return cleaned_codes
        
    except Exception as e:
        st.error(f"Error extracting PO product codes: {e}")
        return []
    
def check_vsba_in_po_line(pdf_file):
    """
    Check if "VSBA" appears in the same line as the PO number in the PO PDF.
    Returns True if VSBA is found in the same line as the PO number, False otherwise.
    """
    try:
        # First, extract the PO number
        po_number = extract_po_number(pdf_file)
        if not po_number:
            return False
        
        # Now, search for the line containing the PO number
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        # Check if this line contains the PO number
                        if po_number in line:
                            # Now check if "VSBA" is in the same line
                            if "VSBA" in line.upper():
                                return True
        
        return False
        
    except Exception as e:
        st.error(f"Error checking VSBA in PO line: {e}")
        return False
    
def extract_item_description_product_code_and_check_vsba(pdf_file):
    """
    Extract the product code from the second line under "Item Description" in the PO PDF.
    Also checks if "VSBA" is present at the end of the product code.
    Returns a tuple: (product_code, vsba_found)
    """
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for i, line in enumerate(lines):
                        # Look for "Item Description" line
                        if "Item Description" in line:
                            # The product code should be in the next line (i+1)
                            if i + 1 < len(lines):
                                next_line = lines[i+1].strip()
                                # Check if this line contains a product code pattern
                                # Example: AG.PRC.TKT_PILB 497_REG_L47.625XW28.575mm-336593-VSBA
                                if next_line and ("PRC.TKT" in next_line or "AG.PRC.TKT" in next_line):
                                    # Check if VSBA is at the end
                                    vsba_found = next_line.upper().endswith("VSBA")
                                    return next_line, vsba_found
        return "", False
    except Exception as e:
        
        st.error(f"Error extracting item description product code: {e}")
        return "", False

def extract_wo_product_code_with_vsba(pdf_file):
    """
    Extract WO product codes and check if they contain VSBA.
    Returns a dictionary with product code and VSBA status.
    """
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        
        lines = text.split("\n")
        wo_codes_with_vsba = []
        
        for line in lines:
            # Look for "Product Code" line
            if "Product Code" in line:
                # Extract everything after "Product Code:"
                code_match = re.search(r"Product Code[:\s]*(.*)", line, re.IGNORECASE)
                if code_match:
                    full_code_line = code_match.group(1).strip()
                    
                    # Check if VSBA is present in the line
                    has_vsba = "VSBA" in full_code_line.upper()
                    
                    # Clean the code (remove extra spaces, etc.)
                    cleaned_code = full_code_line.strip()
                    
                    wo_codes_with_vsba.append({
                        "Full_Code_Line": full_code_line,
                        "Cleaned_Code": cleaned_code,
                        "Has_VSBA": has_vsba,
                        "VSBA_Status": "✅ VSBA Found" if has_vsba else "❌ No VSBA"
                    })
        
        return wo_codes_with_vsba
        
    except Exception as e:
        
        st.error(f"Error extracting WO product code with VSBA: {e}")
        return []


def extract_po_product_code_with_vsba(pdf_file):
    """
    Extract PO product codes from multiple patterns and check if they contain VSBA.
    Returns a dictionary with product code and VSBA status.
    """
    try:
        pdf_file.seek(0)
        with pdfplumber.open(pdf_file) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        
        lines = text.split("\n")
        po_codes_with_vsba = []
        
        # Pattern 1: Look for "Sup. Ref." or "Supplier Reference"
        for line in lines:
            # Sup. Ref. pattern
            sup_ref_match = re.search(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*(.+)", line, re.IGNORECASE)
            if sup_ref_match:
                full_code_line = sup_ref_match.group(1).strip()
                
                # Check if VSBA is present
                has_vsba = "VSBA" in full_code_line.upper()
                
                po_codes_with_vsba.append({
                    "Source": "Supplier Reference",
                    "Full_Code_Line": full_code_line,
                    "Has_VSBA": has_vsba,
                    "VSBA_Status": "✅ VSBA Found" if has_vsba else "❌ No VSBA"
                })
        
        # Pattern 2: Look for TAG.PRC.TKT patterns
        tag_prc_matches = re.findall(r"(TAG\.PRC\.TKT_[^_]+_[^_]+[^\s]*)", text)
        for match in tag_prc_matches:
            full_code_line = match.strip()
            has_vsba = "VSBA" in full_code_line.upper()
            
            po_codes_with_vsba.append({
                "Source": "TAG.PRC.TKT",
                "Full_Code_Line": full_code_line,
                "Has_VSBA": has_vsba,
                "VSBA_Status": "✅ VSBA Found" if has_vsba else "❌ No VSBA"
            })
        
        # Pattern 3: Look for TAG.HANG patterns
        tag_hang_matches = re.findall(r"(TAG\.HANG_[^_]+_[^_]+[^\s]*)", text)
        for match in tag_hang_matches:
            full_code_line = match.strip()
            has_vsba = "VSBA" in full_code_line.upper()
            
            po_codes_with_vsba.append({
                "Source": "TAG.HANG",
                "Full_Code_Line": full_code_line,
                "Has_VSBA": has_vsba,
                "VSBA_Status": "✅ VSBA Found" if has_vsba else "❌ No VSBA"
            })
        
        # Pattern 4: Look in Item Description section
        item_desc_match = re.search(r"Item Description[:\s]*\n(.+)", text, re.IGNORECASE)
        if item_desc_match:
            full_code_line = item_desc_match.group(1).strip()
            has_vsba = "VSBA" in full_code_line.upper()
            
            po_codes_with_vsba.append({
                "Source": "Item Description",
                "Full_Code_Line": full_code_line,
                "Has_VSBA": has_vsba,
                "VSBA_Status": "✅ VSBA Found" if has_vsba else "❌ No VSBA"
            })
        
        # Remove duplicates based on Full_Code_Line
        seen = set()
        unique_codes = []
        for code_info in po_codes_with_vsba:
            if code_info["Full_Code_Line"] not in seen:
                seen.add(code_info["Full_Code_Line"])
                unique_codes.append(code_info)
        
        return unique_codes
        
    except Exception as e:
        
        st.error(f"Error extracting PO product code with VSBA: {e}")
        return []


def compare_vsba_status(wo_vsba_data, po_vsba_data):
    """
    Compare VSBA status between WO and PO product codes.
    Returns a summary dictionary.
    """
    wo_has_vsba = any(item["Has_VSBA"] for item in wo_vsba_data)
    po_has_vsba = any(item["Has_VSBA"] for item in po_vsba_data)
    
    return {
        "WO_VSBA_Found": wo_has_vsba,
        "PO_VSBA_Found": po_has_vsba,
        "Both_Have_VSBA": wo_has_vsba and po_has_vsba,
        "Status": "✅ Both have VSBA" if (wo_has_vsba and po_has_vsba) else 
                  "⚠️ Only WO has VSBA" if wo_has_vsba else 
                  "⚠️ Only PO has VSBA" if po_has_vsba else 
                  "❌ Neither has VSBA"
    }