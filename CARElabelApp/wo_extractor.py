import streamlit as st
import os
import tempfile
import shutil
import sys
import pandas as pd
import pdfplumber
import PyPDF2
import re
from datetime import datetime
from typing import List, Dict, Any

# ----------------- Helper Functions for WO Data Extraction -----------------
@st.cache_data
def extract_text_from_pdf(pdf_file):
    """Extract text from PDF file with caching"""
    try:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() or ""
        return text
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None

# ----------------- Data Extraction Functions for WO Only -----------------
@st.cache_data
def extract_po_number(text):
    """Extract PO Number from WO documents with caching"""
    # Try WO format
    order_match = re.search(r'VS PO Number:\s*(\d+)', text)
    if not order_match:
        order_match = re.search(r'Customer Order No:\s*(\d+)', text)
    
    if order_match:
        return order_match.group(1).strip()
    
    return "Not Found"

@st.cache_data
def extract_season(text):
    """Extract Season from WO documents with caching"""
    # Try WO format first - stop at "Line Item:" but don't include it
    season_match = re.search(r'Season:\s*(.*?)\s*(?=Line Item:|$)', text, re.IGNORECASE | re.DOTALL)
    if season_match:
        season = season_match.group(1).strip()
        # Clean up any extra whitespace
        season = re.sub(r'\s+', ' ', season)
        return season
    
    # Fallback: Try original pattern if "Line Item:" not found
    season_match = re.search(r'Season:\s*([^\n]+)', text)
    if season_match:
        return season_match.group(1).strip()
    
    return "Not Found"

@st.cache_data
def extract_factory_id(text):
    """Extract Factory ID from WO documents with caching"""
    # Try WO format first
    factory_match = re.search(r'Factory ID:\s*(\d+)', text)
    if factory_match:
        return factory_match.group(1).strip()
    
    return "Not Found"

@st.cache_data
def extract_silhouette(text):
    """Extract Silhouette from WO documents with caching"""
    # Try WO format first - Extract only text before the first slash or dash
    silhouette_match = re.search(r'Silhouette:\s*([A-Z\s]+?)(?:/|-)', text)
    if silhouette_match:
        return silhouette_match.group(1).strip()
    
    # Fallback: Try simple description pattern
    silhouette_match = re.search(r'Description:\s*([^\n]+)', text)
    return silhouette_match.group(1).strip() if silhouette_match else "Not Found"

@st.cache_data
def extract_vss_vsd(text):
    """Extract VSS# or VSD# from WO documents with caching"""
    # Try WO format first - specifically from Product Details section
    product_details_match = re.search(r'Product Details:(.*?)(?:Additional Instructions:|Garment Components|End of Works Order)', text, re.DOTALL)
    
    if product_details_match:
        product_details_text = product_details_match.group(1)
        
        # Look for VSS# in the Product Details section
        vss_match = re.search(r'VSS#\s*:\s*(\d+)', product_details_text)
        if vss_match:
            return vss_match.group(1)
        
        # Look for VSD# in the Product Details section
        vsd_match = re.search(r'VSD#\s*:\s*([A-Za-z0-9\-]+)', product_details_text)
        if vsd_match:
            return vsd_match.group(1)
    
    # If not found in Product Details, try the entire document as fallback
    vss_match = re.search(r'VSS#\s*:\s*(\d+)', text)
    if vss_match:
        return vss_match.group(1)
    
    vsd_match = re.search(r'VSD#\s*:\s*([A-Za-z0-9\-]+)', text)
    if vsd_match:
        return vsd_match.group(1)
    
    return "Not Found"

@st.cache_data
def extract_date_of_mfr(text):
    """Extract Date of MFR from WO documents with caching"""
    # Try WO format first
    mfr_match = re.search(r'Date of MFR#:\s*(\d+/\d+)', text)
    if mfr_match:
        return mfr_match.group(1).strip()
    
    return "Not Found"

@st.cache_data
def extract_country_of_origin(text):
    """Extract Country of Origin from WO documents with caching"""
    # Try WO format first - "made in Sri Lanka" (captures multi-word country names)
    coo_match = re.search(r'made in\s+([A-Za-z\s]+?)(?:/|$)', text, re.IGNORECASE)
    if coo_match:
        return coo_match.group(1).strip()
    
    return "Not Found"

@st.cache_data
def extract_size_breakdown(text):
    """Extract Size/Age Breakdown from WO documents with caching"""
    
    # --- WO FORMAT: Specific table under Product Details ---
    # Step 1: Isolate the Product Details section to narrow the search
    product_details_match = re.search(r'Product Details:(.*?)(?:Additional Instructions:|Garment Components|End of Works Order|$)', text, re.DOTALL)
    
    if product_details_match:
        product_details_text = product_details_match.group(1)
        
        # Step 2: Within Product Details, find the Size/Age Breakdown section
        size_breakdown_match = re.search(r'Size/Age Breakdown:(.*?)(?:ITL Factory Code:|Care Instruction Set|$)', product_details_text, re.DOTALL)
        
        if size_breakdown_match:
            size_text = size_breakdown_match.group(1)
            
            # Step 3: Look for the specific table under "Panties/Swim Bottoms"
            # This regex finds the header and captures the table content until the next major section or end of text.
            table_match = re.search(r'Panties/Swim Bottoms\s*[\-:\s]*\n(.*?)(?=\n\n|\n[A-Za-z\s]+\n|$)', size_text, re.DOTALL)
            
            if table_match:
                table_content = table_match.group(1)
                size_breakdown = {}
                
                # Step 4: Parse each row of the table to get size (col 1) and quantity (col 3)
                # This regex processes each line, capturing the size and quantity while ignoring the middle column.
                # It's anchored to the start/end of each line for accuracy.
                row_pattern = r'^\s*([A-Z]{1,4})\s+.*?\s+([\d,]+(?:\.\d+)?)\s*$'
                
                # Find all lines that match the row pattern
                for size, qty in re.findall(row_pattern, table_content, re.MULTILINE):
                    # Clean up quantity by removing commas and convert to float
                    try:
                        clean_qty = float(qty.replace(',', ''))
                        size_breakdown[size] = clean_qty
                    except (ValueError, TypeError):
                        # Skip if the quantity is not a valid number
                        continue
                
                # If we successfully found data in this specific format, return it immediately
                if size_breakdown:
                    return size_breakdown

    # --- FALLBACK 1: Original WO format if the specific table wasn't found ---
    # This part handles other WO formats that are not in a table.
    size_section_match = re.search(r'Size/Age Breakdown:(.*?)(?:ITL Factory Code:|$)', text, re.DOTALL)
    if size_section_match:
        size_text = size_section_match.group(1)
        # Match patterns like "XS | XP | ECH | 165/64A 696"
        size_lines = re.findall(r'([A-Z]+(?:/[A-Z0-9]+)*)\s+(\d+)', size_text)
        size_breakdown = {}
        for size, qty in size_lines:
            # Extract main size (first part before /)
            main_size = size.split('/')[0]
            if main_size in size_breakdown:
                size_breakdown[main_size] += int(qty)
            else:
                size_breakdown[main_size] = int(qty)
        
        if size_breakdown:
            return size_breakdown
    
    # Default empty breakdown if nothing found
    return {}

@st.cache_data
def extract_care_instruction(text):
    """Extract Care Instruction from WO documents with caching"""
    # Try WO format first
    care_match = re.search(r'Care Instruction Set 1:\s*([^\n]+)', text)
    if care_match:
        return care_match.group(1).strip()
    
    return "Not Found"

@st.cache_data
def extract_quantity(text):
    """Extract Quantity from WO documents with caching"""
    # Try WO format first
    quantity_match = re.search(r'Quantity:\s*([\d,]+)', text)
    if quantity_match:
        # Remove commas and convert to integer
        quantity = quantity_match.group(1).replace(',', '')
        return int(quantity)
    
    # Calculate from size breakdown if Quantity is not found
    size_breakdown = extract_size_breakdown(text)
    total_quantity = sum(size_breakdown.values())
    
    if total_quantity > 0:
        if total_quantity.is_integer():
            return int(total_quantity)
        else:
            return total_quantity
    
    return "Not Found"

@st.cache_data
def extract_garment_components(text):
    """Extract Garment Components & Fibre Contents from WO documents with caching"""
    # Try WO format first
    fibre_section = re.search(r'Garment Components &\s*Fibre Contents:(.*?)Care Instructions:', text, re.DOTALL)
    if fibre_section:
        fibre_text = fibre_section.group(1).strip()
        # Clean up extra whitespace
        fibre_text = re.sub(r'\s+', ' ', fibre_text)
        return fibre_text
    
    return "Not Found"

# New extraction functions for the additional fields
@st.cache_data
def extract_product_code(text):
    """Extract Product Code from WO documents with caching - read only the value after 'Product Code:' within Product Details section"""
    # Step 1: Find the Product Details section
    product_details_match = re.search(r'Product Details:(.*?)(?:Additional Instructions:|Garment Components|End of Works Order|$)', text, re.DOTALL)
    
    if product_details_match:
        # Step 2: Extract only the content within Product Details section
        product_details_text = product_details_match.group(1)
        
        # Step 3: Within Product Details section, find the Product Code line
        # Pattern 1: Look for "Product Code:" followed by typical product code format (letters-numbers-spaces-slashes)
        product_code_match = re.search(r'Product Code:\s*([A-Za-z0-9\-/\s]+?)(?=\s*(?:Product Description:|Quantity:|$))', product_details_text, re.IGNORECASE)
        if product_code_match:
            # Step 4: Extract only the value after "Product Code:"
            product_code_value = product_code_match.group(1).strip()
            # Step 5: Clean up extra whitespace
            product_code_value = re.sub(r'\s+', ' ', product_code_value)
            return product_code_value
        
        # Pattern 2: Fallback - try a more general pattern but limit the length
        product_code_match = re.search(r'Product Code:\s*([^\n]{1,50})', product_details_text, re.IGNORECASE)
        if product_code_match:
            product_code_value = product_code_match.group(1).strip()
            # Clean up extra whitespace
            product_code_value = re.sub(r'\s+', ' ', product_code_value)
            return product_code_value
    
    return "Not Found"

@st.cache_data
def extract_delivery_date(text):
    """Extract Delivery Date from WO documents with caching"""
    # Try WO format first - Delivery Date (ex factory) under Order Delivery Details
    delivery_date_match = re.search(r"Delivery Date \(ex factory\)\s*:\s*(\d{4}[/\-]\d{2}[/\-]\d{2})", text, re.IGNORECASE)
    if delivery_date_match:
        return delivery_date_match.group(1).strip()
    
    return "Not Found"

@st.cache_data
def extract_delivery_location(text):
    """Extract Customer Delivery Name + Delivery To from WO documents with caching"""
    # Try WO format first - Look for Order Delivery Details section
    # Extract Customer Delivery Name and Deliver To
    customer_name_match = re.search(r'Customer Delivery\s*Name\s*:\s*([^\n]+)', text, re.IGNORECASE)
    deliver_to_match = re.search(r'Deliver To\s*:\s*([^\n]+)', text, re.IGNORECASE)
    
    if customer_name_match and deliver_to_match:
        customer_name = customer_name_match.group(1).strip()
        deliver_to = deliver_to_match.group(1).strip()
        
        # Remove "BFF-" prefix from both customer name and deliver to
        customer_name = re.sub(r'\bBFF-\s*', '', customer_name, flags=re.IGNORECASE)
        deliver_to = re.sub(r'\bBFF-\s*', '', deliver_to, flags=re.IGNORECASE)
        
        # Check if "Sri Lanka" is in the deliver_to address
        if 'sri lanka' in deliver_to.lower():
            # If "Sri Lanka" is found, extract only the address part up to "Sri Lanka"
            # and explicitly skip everything after "Sri Lanka"
            sri_lanka_match = re.search(r'(.*?)\s*Sri Lanka', deliver_to, re.IGNORECASE)
            if sri_lanka_match:
                deliver_to = sri_lanka_match.group(1).strip()
            else:
                # If for some reason the match fails, set deliver_to to empty
                deliver_to = ""
        
        # Combine Customer Delivery Name + Deliver To
        # Check if customer_name is already contained in deliver_to to avoid duplication
        if customer_name.lower() in deliver_to.lower() or deliver_to.lower() in customer_name.lower():
            # If one is contained in the other, use only the longer one
            delivery_location = customer_name if len(customer_name) > len(deliver_to) else deliver_to
        else:
            # If they're different, combine them
            delivery_location = f"{customer_name} + {deliver_to}"
        
        # Clean up multiple commas and spaces
        delivery_location = re.sub(r',\s*,+', ',', delivery_location)  # Remove duplicate commas
        delivery_location = re.sub(r'\s+', ' ', delivery_location)  # Normalize spaces
        
        return delivery_location.strip()
    
    return "Not Found"

@st.cache_data
def extract_color_code(text):
    """Extract Color Code from WO documents with caching"""
    # Try WO format first - from Product Details section
    # Pattern: Product Code: PWLB-165 C/1 / LB 5792, extract "C/1"
    product_code_match = re.search(r'Product Code:\s*[A-Za-z0-9\-]+\s+([A-Za-z0-9/]+)', text)
    if product_code_match:
        color_code = product_code_match.group(1).strip()
        # For example, from "C/1 / LB 5792", extract "C/1"
        color_parts = color_code.split('/')
        if len(color_parts) >= 2:
            return f"{color_parts[0]}/{color_parts[1]}"
        else:
            return color_parts[0]
    
    return "Not Found"

@st.cache_data
def extract_size_id(text):
    """Extract Size ID from WO documents with caching - read only the value after 'Size ID:' within Product Details section"""
    # Step 1: Find the Product Details section
    product_details_match = re.search(r'Product Details:(.*?)(?:Additional Instructions:|Garment Components|End of Works Order|$)', text, re.DOTALL)
    
    if product_details_match:
        # Step 2: Extract only the content within Product Details section
        product_details_text = product_details_match.group(1)
        
        # Step 3: Within Product Details section, find the Size ID line
        # Pattern 1: Look for "Size ID:" followed by typical size ID format
        size_id_match = re.search(r'Size ID:\s*([A-Za-z0-9\-/\s]+?)(?=\s*(?:Product Code:|Product Description:|Quantity:|Size/Age Breakdown:|$))', product_details_text, re.IGNORECASE)
        if size_id_match:
            # Step 4: Extract only the value after "Size ID:"
            size_id_value = size_id_match.group(1).strip()
            # Step 5: Clean up extra whitespace
            size_id_value = re.sub(r'\s+', ' ', size_id_value)
            # Step 6: Remove "Size/Age Breakdown" if it appears in the value
            size_id_value = re.sub(r'Size/Age\s*Breakdown.*$', '', size_id_value, flags=re.IGNORECASE).strip()
            return size_id_value
        
        # Pattern 2: Fallback - try a more general pattern but limit the length
        size_id_match = re.search(r'Size ID:\s*([^\n]{1,50})', product_details_text, re.IGNORECASE)
        if size_id_match:
            size_id_value = size_id_match.group(1).strip()
            # Clean up extra whitespace
            size_id_value = re.sub(r'\s+', ' ', size_id_value)
            # Remove "Size/Age Breakdown" if it appears in the value
            size_id_value = re.sub(r'Size/Age\s*Breakdown.*$', '', size_id_value, flags=re.IGNORECASE).strip()
            return size_id_value
    
    return "Not Found"

# NEW EXTRACTION FUNCTIONS FOR ADDRESS AND CUSTOMER
@st.cache_data
def extract_address(text):
    """Extract Address from WO documents - under 'works order No:' and in front of 'To:' with caching"""
    # Look for pattern: works order No: [some text] To: [address] Customer:
    # The address is between "To:" and "Customer:"
    address_match = re.search(r'works order No:.*?To:\s*(.*?)\s*Customer:', text, re.IGNORECASE | re.DOTALL)
    if address_match:
        address = address_match.group(1).strip()
        # Clean up extra whitespace and newlines
        address = re.sub(r'\s+', ' ', address)
        return address
    
    # Fallback: Try to find address after "To:" without "Customer:" requirement
    address_match = re.search(r'To:\s*([^\n]+)', text, re.IGNORECASE)
    if address_match:
        address = address_match.group(1).strip()
        # Clean up extra whitespace
        address = re.sub(r'\s+', ' ', address)
        return address
    
    return "Not Found"

@st.cache_data
def extract_customer(text):
    """Extract Customer from WO documents - after 'customer:' and read the infront of line with caching"""
    # Look for pattern: Customer: [customer name]
    customer_match = re.search(r'Customer:\s*([^\n]+)', text, re.IGNORECASE)
    if customer_match:
        customer = customer_match.group(1).strip()
        # Clean up extra whitespace
        customer = re.sub(r'\s+', ' ', customer)
        return customer
    
    return "Not Found"

# Add this function to clean up the season value after extraction
def clean_season_value(season_value):
    """Remove 'Line Item:' and everything after it from the season value"""
    if season_value and season_value != "Not Found":
        # Check if "Line Item:" is in the season value
        if "Line Item:" in season_value:
            # Split at "Line Item:" and take only the part before it
            season_value = season_value.split("Line Item:")[0].strip()
    
    return season_value

# Add this function to clean up the customer value after extraction
def clean_customer_value(customer_value):
    """Remove 'Customer Order No' and everything after it from the customer value"""
    if customer_value and customer_value != "Not Found":
        # Check if "Customer Order No" is in the customer value
        if "Customer Order No" in customer_value:
            # Split at "Customer Order No" and take only the part before it
            customer_value = customer_value.split("Customer Order No")[0].strip()
    
    return customer_value

# ============= NEW: IMPROVED SIZE/AGE BREAKDOWN TABLE EXTRACTION =============
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

def clean_quantity(qty_str):
    """Convert strings like '1,148.0000' or '465.0000' into floats with 4 decimal places"""
    if not qty_str:
        return 0.0000
    qty_str = str(qty_str).strip().replace(",", "")
    try:
        return round(float(qty_str), 4)
    except ValueError:
        return 0.0000

def extract_wo_items_table_enhanced(pdf_file, product_codes=None):
    """
    Enhanced function to extract WO items from Victoria's Secret price ticket tables
    with improved table detection and data extraction for all formats, including sizes split across lines
    Removed fields: Colour, Retail (US), Retail (CA), Multi Price, SKU, Article
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
                    if any(term in row_text for term in ["Style", "Size", "Quantity"]):
                        header_row_idx = i
                        
                        # Map column positions - REMOVED colour, retail, multi, sku, article
                        for j, cell in enumerate(row):
                            cell_text = str(cell).strip().lower() if cell else ""
                            if "style" in cell_text:
                                column_positions["style"] = j
                            elif "size 1" in cell_text or "size" in cell_text:
                                column_positions["size1"] = j
                            elif "size 2" in cell_text:
                                column_positions["size2"] = j
                            elif "panty" in cell_text:
                                column_positions["panty_length"] = j
                            elif "quantity" in cell_text or "qty" in cell_text:
                                column_positions["quantity"] = j
                        break
                
                # If we couldn't find a header row, try to infer it
                if header_row_idx == -1:
                    for i, row in enumerate(processed_table):
                        if not row or len(row) < 3:  # Reduced minimum columns
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
                                    "size1": 1,  # Adjusted positions
                                    "size2": 2,
                                    "panty_length": 3,
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
                        
                        # Extract size1 with special handling for multi-line cells
                        size1_raw = str(row[column_positions.get("size1", 1)] or "")
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
                size = match.group(3)
                quantity = match.group(4).replace(',', '')
                
                try:
                    items.append({
                        "Style": style,
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
                                    "Size 1": size,
                                    "Quantity": int(quantity.replace(',', '')),
                                    "WO Product Code": " / ".join(product_codes) if product_codes else ""
                                })
                            except ValueError:
                                pass
                    
                    i += 1
                
    # Aggregate quantities for items with same style and size
    aggregated_items = {}
    for item in items:
        key = (item["Style"], item["Size 1"])
        if key in aggregated_items:
            aggregated_items[key]["Quantity"] += item["Quantity"]
        else:
            aggregated_items[key] = item.copy()
    
    # Convert back to list
    final_items = list(aggregated_items.values())
    
    return final_items

# ============= NEW: ROBUST SIZE BREAKDOWN TABLE EXTRACTION =============
def extract_size_breakdown_table_from_pdf_robust(pdf_file) -> List[Dict[str, Any]]:
    """
    Robust extraction of Size/Age Breakdown table from WO PDF.
    Prioritizes direct table extraction over text parsing.
    
    Args:
        pdf_file: A file-like object (e.g., from st.file_uploader)
    
    Returns:
        List of dictionaries with 'Size' and 'Order Quantity' keys
    """
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                # Method 1: Try standard table extraction first
                tables = page.extract_tables()
                
                if tables:
                    for table in tables:
                        # Try to find the size breakdown table within the extracted tables
                        size_data = _process_table_for_size_breakdown(table)
                        if size_data:
                            return size_data
                
                # Method 2: If standard fails, try with explicit lines strategy
                tables_explicit = page.extract_tables({
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "explicit_vertical_lines": page.curves + page.edges,
                    "explicit_horizontal_lines": page.curves + page.edges,
                })
                
                if tables_explicit:
                    for table in tables_explicit:
                        size_data = _process_table_for_size_breakdown(table)
                        if size_data:
                            return size_data

                # Method 3: If table extraction fails, try text-based extraction as a last resort
                text = page.extract_text() or ""
                if "Size/Age Breakdown:" in text:
                    size_data = _extract_size_breakdown_from_text(text)
                    if size_data:
                        return size_data

    except Exception as e:
        st.error(f"Error in robust PDF extraction: {e}")
    
    return []

def _process_table_for_size_breakdown(table: List[List[Any]]) -> List[Dict[str, Any]]:
    """
    Process an extracted table to find and extract size breakdown data.
    """
    if not table or len(table) < 2:
        return []

    # Convert all cells to strings and strip whitespace
    processed_table = []
    for row in table:
        processed_row = [str(cell).strip() if cell is not None else "" for cell in row]
        processed_table.append(processed_row)

    # Try to identify the header row
    header_row_idx = _find_header_row(processed_table)
    if header_row_idx == -1:
        # If no clear header, try to infer based on content
        header_row_idx = _infer_header_row(processed_table)
    
    if header_row_idx == -1:
        return []

    # Identify column indices
    size_col_idx, qty_col_idx = _identify_size_quantity_columns(processed_table, header_row_idx)
    
    if size_col_idx == -1 or qty_col_idx == -1:
        return []

    # Extract data from rows below the header
    extracted_data = []
    for row in processed_table[header_row_idx + 1:]:
        if not row or len(row) <= max(size_col_idx, qty_col_idx):
            continue
        
        size = row[size_col_idx].strip()
        qty_str = row[qty_col_idx].strip()
        
        # Clean the size value
        size = _clean_size_value(size)
        
        # Clean the quantity value
        quantity = _clean_quantity_value(qty_str)
        
        if size and quantity > 0:
            extracted_data.append({
                "Size": size,
                "Order Quantity": quantity
            })

    return extracted_data

def _find_header_row(table: List[List[str]]) -> int:
    """
    Find the header row by looking for keywords.
    """
    for i, row in enumerate(table):
        row_text = " ".join(row).upper()
        if any(keyword in row_text for keyword in [
            "SIZE", "AGE", "BREAKDOWN", "QUANTITY", "QTY", "ORDER"
        ]):
            return i
    return -1

def _infer_header_row(table: List[List[str]]) -> int:
    """
    Infer the header row by looking for data patterns.
    """
    for i, row in enumerate(table):
        if len(row) >= 2:
            # Check if first column looks like a size (XS, S, M, L, etc.)
            first_cell = row[0].upper()
            if first_cell in ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]:
                # This might be a data row, so header is likely above
                return max(0, i - 1)
            
            # Check if any cell contains size-like data
            for cell in row:
                cell_upper = cell.upper()
                if any(size in cell_upper for size in ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]):
                    return max(0, i - 1)
    return 0  # Default to first row if nothing found

def _identify_size_quantity_columns(table: List[List[str]], header_row_idx: int) -> tuple:
    """
    Identify which columns contain size and quantity data.
    """
    if header_row_idx >= len(table):
        return -1, -1
    
    header_row = table[header_row_idx]
    size_col_idx = -1
    qty_col_idx = -1
    
    # Look for column headers
    for i, header in enumerate(header_row):
        header_upper = header.upper()
        
        if "SIZE" in header_upper and size_col_idx == -1:
            size_col_idx = i
        elif ("QUANTITY" in header_upper or "QTY" in header_upper) and qty_col_idx == -1:
            qty_col_idx = i
    
    # If headers not found, try to infer from data patterns
    if size_col_idx == -1 or qty_col_idx == -1:
        size_col_idx, qty_col_idx = _infer_columns_from_data(table, header_row_idx)
    
    return size_col_idx, qty_col_idx

def _infer_columns_from_data(table: List[List[str]], header_row_idx: int) -> tuple:
    """
    Infer column indices by analyzing data patterns in rows.
    """
    size_col_idx = -1
    qty_col_idx = -1
    
    # Check rows below header
    for row in table[header_row_idx + 1:header_row_idx + 5]:  # Check next 5 rows
        if len(row) < 2:
            continue
        
        for i, cell in enumerate(row):
            cell_upper = cell.upper()
            
            # Check for size patterns
            if size_col_idx == -1 and cell_upper in ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]:
                size_col_idx = i
            
            # Check for quantity patterns (numbers)
            if qty_col_idx == -1 and re.match(r'^[\d,]+$', cell):
                qty_col_idx = i
        
        # If both found, break
        if size_col_idx != -1 and qty_col_idx != -1:
            break
    
    return size_col_idx, qty_col_idx

def _clean_size_value(size_str: str) -> str:
    """
    Clean and normalize size value.
    """
    if not size_str:
        return ""
    
    # Remove common separators and take the first part
    size_str = size_str.split('|')[0].split('/')[0].strip()
    
    # Ensure it's a valid size
    valid_sizes = ["XS", "S", "M", "L", "XL", "XXL", "XXXL", "P", "G", "XG"]
    if size_str.upper() in valid_sizes:
        return size_str.upper()
    
    # If not a standard size, return as-is but cleaned
    return size_str.strip()

def _clean_quantity_value(qty_str: str) -> int:
    """
    Clean and convert quantity value to integer.
    """
    if not qty_str:
        return 0
    
    # Remove commas and any non-digit characters
    qty_clean = re.sub(r'[^\d]', '', qty_str)
    
    try:
        return int(qty_clean)
    except ValueError:
        return 0

def _extract_size_breakdown_from_text(text: str) -> List[Dict[str, Any]]:
    """
    Extract size breakdown from text as a fallback method.
    """
    # Find the Size/Age Breakdown section
    size_breakdown_match = re.search(
        r'Size/Age\s*Breakdown:\s*(.*?)(?=ITL\s*Factory\s*Code:|VSD#:|VSS#:|RN#:|CA#:|Factory\s*ID:|Product\s*Details:|$)',
        text,
        re.IGNORECASE | re.DOTALL
    )
    
    if not size_breakdown_match:
        return []
    
    size_section = size_breakdown_match.group(1).strip()
    lines = [line.strip() for line in size_section.split('\n') if line.strip()]
    
    # Skip header line if present
    if lines and any(header in lines[0].upper() for header in ['ORDER QUANTITY', 'QUANTITY', 'QTY']):
        lines = lines[1:]
    
    extracted_data = []
    for line in lines:
        # Try to match size and quantity patterns
        # Pattern: SIZE ... QUANTITY
        match = re.match(r'^([A-Z]{1,4})\s+.*?(\d{1,3}(?:,\d{3})*)$', line)
        if match:
            size = match.group(1).strip()
            quantity = int(match.group(2).replace(',', ''))
            
            if size and quantity > 0:
                extracted_data.append({
                    "Size": size,
                    "Order Quantity": quantity
                })
    
    return extracted_data

def extract_size_breakdown_table_robust(pdf_file):
    """
    Main function to call the robust size breakdown extraction.
    This is the function you should call from your main application.
    """
    return extract_size_breakdown_table_from_pdf_robust(pdf_file)

# Keep the original function for backward compatibility
def extract_size_breakdown_table(text):
    """
    Extract the complete Size/Age Breakdown table from WO documents.
    This function specifically looks for the table after "Size/Age Breakdown:" 
    and extracts size and order quantity data until "ITL Factory Code:".
    Returns a list of dictionaries with Size and Order Quantity.
    """
    # Find the Size/Age Breakdown section
    size_breakdown_match = re.search(
        r'Size/Age Breakdown:\s*(.*?)(?=ITL Factory Code:|VSD#:|VSS#:|RN#:|CA#:|Factory ID:|$)',
        text,
        re.DOTALL
    )
    
    if not size_breakdown_match:
        return []
    
    size_section = size_breakdown_match.group(1).strip()
    
    # Split into lines and filter out empty lines
    lines = [line.strip() for line in size_section.split('\n') if line.strip()]
    
    # The first line might be a header, check if it contains "Order Quantity"
    if lines and ('Order Quantity' in lines[0] or 'Apparel/Sleep Tops' in lines[0]):
        lines = lines[1:]  # Skip the header line
    
    extracted_rows = []
    
    # Process each line to extract size and quantity
    for line in lines:
        # Skip lines that don't contain a size pattern
        if not re.search(r'^[A-Z]{1,4}(\s*\|\s*[A-Z]{1,4})*', line):
            continue
        
        # Extract the size (first element before the first pipe or space)
        size_match = re.match(r'^([A-Z]{1,4})', line)
        if not size_match:
            continue
        
        size = size_match.group(1)
        
        # Extract the quantity (last number in the line)
        quantity_match = re.search(r'(\d+)\s*$', line)
        if not quantity_match:
            continue
        
        quantity = int(quantity_match.group(1))
        
        # Add to our results
        extracted_rows.append({
            "Size": size,
            "Order Quantity": quantity
        })
    
    return extracted_rows

# ----------------- Main Parsing Functions -----------------
@st.cache_data
def parse_wo_data(text):
    """Parse Work Order data using extraction functions with caching"""
    data = {}

    # Extract all data points using functions
    data['PO Number'] = extract_po_number(text)
    data['Season'] = clean_season_value(extract_season(text))  # Clean the season value
    data['Factory ID'] = extract_factory_id(text)
    data['Silhouette'] = extract_silhouette(text)
    data['VSS/VSD'] = extract_vss_vsd(text)
    data['Date of MFR'] = extract_date_of_mfr(text)
    data['Country of Origin'] = extract_country_of_origin(text)
    
    # NEW: Extract the complete Size/Age Breakdown table
    data['Size Breakdown Table'] = extract_size_breakdown_table(text)
    
    # Keep the original size breakdown for backward compatibility
    data['Size Breakdown'] = extract_size_breakdown(text)
    
    data['Care Instruction'] = extract_care_instruction(text)
    data['Quantity'] = extract_quantity(text)
    data['Garment Components'] = extract_garment_components(text)
    
    # New fields
    data['Product Code'] = extract_product_code(text)
    data['Delivery Location'] = extract_delivery_location(text)
    data['Color Code'] = extract_color_code(text)
    data['Size ID'] = extract_size_id(text)
    data['Delivery Date'] = extract_delivery_date(text)
    
    # NEW FIELDS: Address and Customer
    data['Address'] = extract_address(text)
    data['Customer'] = clean_customer_value(extract_customer(text))

    return data

def process_wo_file(wo_file):
    """Process WO file and return parsed data"""
    # Process WO
    with st.spinner("Processing WO..."):
        wo_text = extract_text_from_pdf(wo_file)
        if wo_text:
            return parse_wo_data(wo_text)
    return None