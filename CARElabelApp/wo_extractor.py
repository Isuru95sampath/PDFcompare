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

def extract_po_number(text):
    """Extract PO Number from WO documents with caching"""
    # Try WO format
    order_match = re.search(r'VS PO Number:\s*(\d+)', text)
    if not order_match:
        order_match = re.search(r'Customer Order No:\s*(\d+)', text)
    
    if order_match:
        return order_match.group(1).strip()
    
    return "Not Found"


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


def extract_factory_id(text):
    """Extract Factory ID from WO documents with caching"""
    # Try WO format first
    factory_match = re.search(r'Factory ID:\s*(\d+)', text)
    if factory_match:
        return factory_match.group(1).strip()
    
    return "Not Found"


def extract_silhouette(text):
    """Extract Silhouette from WO documents with caching"""
    # Try WO format first - Extract only text before the first slash or dash
    silhouette_match = re.search(r'Silhouette:\s*([A-Z\s]+?)(?:/|-)', text)
    if silhouette_match:
        return silhouette_match.group(1).strip()
    
    # Fallback: Try simple description pattern
    silhouette_match = re.search(r'Description:\s*([^\n]+)', text)
    return silhouette_match.group(1).strip() if silhouette_match else "Not Found"


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


def extract_date_of_mfr(text):
    """Extract Date of MFR from WO documents with caching"""
    # Try WO format first
    mfr_match = re.search(r'Date of MFR#:\s*(\d+/\d+)', text)
    if mfr_match:
        return mfr_match.group(1).strip()
    
    return "Not Found"


def extract_country_of_origin(text):
    """Extract Country of Origin from WO documents with caching"""
    # Try WO format first - "made in Sri Lanka" (captures multi-word country names)
    coo_match = re.search(r'made in\s+([A-Za-z\s]+?)(?:/|$)', text, re.IGNORECASE)
    if coo_match:
        return coo_match.group(1).strip()
    
    return "Not Found"


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


def extract_care_instruction(text):
    """Extract Care Instruction from WO documents with caching"""
    # Try WO format first
    care_match = re.search(r'Care Instruction Set 1:\s*([^\n]+)', text)
    if care_match:
        return care_match.group(1).strip()
    
    return "Not Found"


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

def extract_delivery_date(text):
    """Extract Delivery Date from WO documents - handles multiple PDF formats"""
    # Method 1: Look for "Order Delivery Details:" section first, then find the date
    order_delivery_match = re.search(
        r'Order\s+Delivery\s+Details:(.*?)(?=Product\s+Details:|$)',
        text,
        re.IGNORECASE | re.DOTALL
    )
    
    if order_delivery_match:
        delivery_section = order_delivery_match.group(1)
        
        # Pattern 1: Handle line break between "ex" and "factory"
        delivery_date_match = re.search(
            r'Delivery\s+Date\s*\(\s*ex\s*\n\s*factory\s*\)\s*:\s*(\d{4}[/\-]\d{2}[/\-]\d{2})',
            delivery_section,
            re.IGNORECASE
        )
        
        if delivery_date_match:
            return delivery_date_match.group(1).strip()
        
        # Pattern 2: Handle no line break between "ex" and "factory"
        delivery_date_match = re.search(
            r'Delivery\s+Date\s*\(\s*ex\s+factory\s*\)\s*:\s*(\d{4}[/\-]\d{2}[/\-]\d{2})',
            delivery_section,
            re.IGNORECASE
        )
        
        if delivery_date_match:
            return delivery_date_match.group(1).strip()
        
        # Pattern 3: More flexible pattern with any content between parentheses
        delivery_date_match = re.search(
            r'Delivery\s+Date\s*\([^)]*\)\s*:\s*(\d{4}[/\-]\d{2}[/\-]\d{2})',
            delivery_section,
            re.IGNORECASE | re.DOTALL
        )
        
        if delivery_date_match:
            return delivery_date_match.group(1).strip()
        
        # Pattern 4: Even simpler - just find date after "Delivery Date"
        delivery_date_match = re.search(
            r'Delivery\s+Date[^:]*:\s*(\d{4}[/\-]\d{2}[/\-]\d{2})',
            delivery_section,
            re.IGNORECASE
        )
        
        if delivery_date_match:
            return delivery_date_match.group(1).strip()
    
    # Method 2: Search the entire document if the section approach didn't work
    # Pattern 1: Handle line break between "ex" and "factory"
    delivery_date_match = re.search(
        r'Delivery\s+Date\s*\(\s*ex\s*\n\s*factory\s*\)\s*:\s*(\d{4}[/\-]\d{2}[/\-]\d{2})',
        text,
        re.IGNORECASE
    )
    
    if delivery_date_match:
        return delivery_date_match.group(1).strip()
    
    # Pattern 2: Handle no line break between "ex" and "factory"
    delivery_date_match = re.search(
        r'Delivery\s+Date\s*\(\s*ex\s+factory\s*\)\s*:\s*(\d{4}[/\-]\d{2}[/\-]\d{2})',
        text,
        re.IGNORECASE
    )
    
    if delivery_date_match:
        return delivery_date_match.group(1).strip()
    
    # Pattern 3: More flexible pattern with any content between parentheses
    delivery_date_match = re.search(
        r'Delivery\s+Date\s*\([^)]*\)\s*:\s*(\d{4}[/\-]\d{2}[/\-]\d{2})',
        text,
        re.IGNORECASE | re.DOTALL
    )
    
    if delivery_date_match:
        return delivery_date_match.group(1).strip()
    
    # Pattern 4: Even simpler - just find date after "Delivery Date"
    delivery_date_match = re.search(
        r'Delivery\s+Date[^:]*:\s*(\d{4}[/\-]\d{2}[/\-]\d{2})',
        text,
        re.IGNORECASE
    )
    
    if delivery_date_match:
        return delivery_date_match.group(1).strip()
    
    # Method 3: Try to find any date in YYYY/MM/DD format in the document
    delivery_date_match = re.search(
        r'(\d{4}[/\-]\d{2}[/\-]\d{2})',
        text
    )
    
    if delivery_date_match:
        return delivery_date_match.group(1).strip()
    
    return "Not Found"


def extract_size_id(text):
    """Extract Size ID from WO documents - handles multiple PDF formats"""
    # Method 1: Look for "Size ID:" within the Product Details section
    product_details_match = re.search(
        r'Product\s+Details:(.*?)(?:Additional\s+Instructions:|Garment\s+Components|End\s+of\s+Works\s+Order|$)',
        text,
        re.IGNORECASE | re.DOTALL
    )
    
    if product_details_match:
        product_details_text = product_details_match.group(1)
        
        # Pattern 1: Standard "Size ID:" followed by content
        size_id_match = re.search(
            r'Size\s+ID\s*:\s*([A-Za-z0-9\-/\s]+?)(?=\s*(?:Product\s+Code:|Product\s+Description:|Quantity:|Size/Age\s+Breakdown:|$))',
            product_details_text,
            re.IGNORECASE
        )
        
        if size_id_match:
            size_id_value = size_id_match.group(1).strip()
            size_id_value = re.sub(r'\s+', ' ', size_id_value)
            size_id_value = re.sub(r'Size/Age\s*Breakdown.*$', '', size_id_value, flags=re.IGNORECASE).strip()
            return size_id_value
        
        # Pattern 2: More general pattern
        size_id_match = re.search(
            r'Size\s+ID\s*:\s*([^\n]{1,50})',
            product_details_text,
            re.IGNORECASE
        )
        
        if size_id_match:
            size_id_value = size_id_match.group(1).strip()
            size_id_value = re.sub(r'\s+', ' ', size_id_value)
            size_id_value = re.sub(r'Size/Age\s*Breakdown.*$', '', size_id_value, flags=re.IGNORECASE).strip()
            return size_id_value
    
    # Method 2: Search the entire document if the section approach didn't work
    # Pattern 1: Standard "Size ID:" followed by content
    size_id_match = re.search(
        r'Size\s+ID\s*:\s*([A-Za-z0-9\-/\s]+?)(?=\s*(?:Product\s+Code:|Product\s+Description:|Quantity:|Size/Age\s+Breakdown:|$))',
        text,
        re.IGNORECASE
    )
    
    if size_id_match:
        size_id_value = size_id_match.group(1).strip()
        size_id_value = re.sub(r'\s+', ' ', size_id_value)
        size_id_value = re.sub(r'Size/Age\s*Breakdown.*$', '', size_id_value, flags=re.IGNORECASE).strip()
        return size_id_value
    
    # Pattern 2: More general pattern
    size_id_match = re.search(
        r'Size\s+ID\s*:\s*([^\n]{1,50})',
        text,
        re.IGNORECASE
    )
    
    if size_id_match:
        size_id_value = size_id_match.group(1).strip()
        size_id_value = re.sub(r'\s+', ' ', size_id_value)
        size_id_value = re.sub(r'Size/Age\s*Breakdown.*$', '', size_id_value, flags=re.IGNORECASE).strip()
        return size_id_value
    
    # Method 3: Try to find any size ID pattern in the document
    # Look for patterns like "VSGLOBAL004" or similar size ID formats
    size_id_match = re.search(
        r'(?:Size\s+ID\s*:|ID\s*:)\s*([A-Za-z0-9\-/\s]+)',
        text,
        re.IGNORECASE
    )
    
    if size_id_match:
        size_id_value = size_id_match.group(1).strip()
        size_id_value = re.sub(r'\s+', ' ', size_id_value)
        return size_id_value
    
    # Method 4: Try to find any pattern that looks like a size ID
    # This is a last resort and might not be accurate
    size_id_match = re.search(
        r'([A-Z]{3,}[A-Z0-9]{3,})',
        text
    )
    
    if size_id_match:
        return size_id_match.group(1).strip()
    
    return "Not Found"


def extract_delivery_location(text):
    """Extract Delivery Location from WO documents - combines 'Customer Delivery Name:' and 'Deliver To:' under 'Order Delivery Details:'"""
    # First, find the Order Delivery Details section
    order_delivery_match = re.search(
        r'Order\s+Delivery\s+Details:(.*?)(?=Product\s+Details:|$)',
        text,
        re.IGNORECASE | re.DOTALL
    )
    
    if not order_delivery_match:
        return "Not Found"
    
    delivery_section = order_delivery_match.group(1)
    
    # Extract Customer Delivery Name (handle line breaks)
    customer_name = ""
    customer_name_match = re.search(
        r'Customer\s+Delivery\s*\n\s*Name\s*:\s*([^\n]+)',
        delivery_section,
        re.IGNORECASE
    )
    
    if not customer_name_match:
        # Try without line break
        customer_name_match = re.search(
            r'Customer\s+Delivery\s+Name\s*:\s*([^\n]+)',
            delivery_section,
            re.IGNORECASE
        )
    
    if customer_name_match:
        customer_name = customer_name_match.group(1).strip()
    
    # Extract Deliver To address - multiple approaches
    deliver_to = ""
    
    # Approach 1: Look for "Deliver To:" and capture until the next field that starts with capital letter
    deliver_to_match = re.search(
        r'Deliver\s+To\s*:\s*(.*?)(?=\n\s*[A-Z][a-zA-Z]*\s*:|\n\s*Delivery\s+Method:|\n\s*Delivery\s+Account|\n\s*Contact:|\n\s*Comments|\n\s*Special|\n\s*Product\s+Details:|\Z)',
        delivery_section,
        re.IGNORECASE | re.DOTALL
    )
    
    if deliver_to_match:
        deliver_to = deliver_to_match.group(1).strip()
        # Clean up extra whitespace but preserve the address structure
        deliver_to = re.sub(r'\s+', ' ', deliver_to)
    else:
        # Approach 2: Look for "Deliver To:" and capture everything until next major section
        deliver_to_match = re.search(
            r'Deliver\s+To\s*:\s*(.*?)(?=Delivery\s+Method|Product\s+Details|End\s+of\s+Works\s+Order|$)',
            delivery_section,
            re.IGNORECASE | re.DOTALL
        )
        
        if deliver_to_match:
            deliver_to = deliver_to_match.group(1).strip()
            # Clean up extra whitespace
            deliver_to = re.sub(r'\s+', ' ', deliver_to)
        else:
            # Approach 3: Try to capture just the first line after "Deliver To:"
            deliver_to_match = re.search(
                r'Deliver\s+To\s*:\s*([^\n]+)',
                delivery_section,
                re.IGNORECASE
            )
            
            if deliver_to_match:
                deliver_to = deliver_to_match.group(1).strip()
    
    # Clean up the extracted values
    customer_name = re.sub(r'\bBFF-\s*', '', customer_name, flags=re.IGNORECASE)
    deliver_to = re.sub(r'\bBFF-\s*', '', deliver_to, flags=re.IGNORECASE)
    
    # Combine the two parts
    if customer_name and deliver_to:
        # Check if customer_name is already contained in deliver_to to avoid duplication
        if customer_name.lower() in deliver_to.lower() or deliver_to.lower() in customer_name.lower():
            # If one is contained in the other, use only the longer one
            delivery_location = customer_name if len(customer_name) > len(deliver_to) else deliver_to
        else:
            # If they're different, combine them
            delivery_location = f"{customer_name} - {deliver_to}"
    elif customer_name:
        delivery_location = customer_name
    elif deliver_to:
        delivery_location = deliver_to
    else:
        delivery_location = "Not Found"
    
    # Clean up final result
    delivery_location = re.sub(r'\s+', ' ', delivery_location).strip()
    
    return delivery_location

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
def extract_size_breakdown_table_robust(pdf_file):
    """
    Main function to call the robust size breakdown extraction.
    This is the function you should call from your main application.
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
    Ensures 'Order Quantity' is converted to a numeric type.
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
        return []

    # Get the header row to determine column positions
    header_row = processed_table[header_row_idx]
    
    # Find the columns we need
    size_col = None
    quantity_col = None
    
    for i, cell in enumerate(header_row):
        cell_lower = cell.lower()
        if "size" in cell_lower and size_col is None:
            size_col = i
        elif ("quantity" in cell_lower or "qty" in cell_lower) and quantity_col is None:
            quantity_col = i
    
    # If we couldn't find the columns, try to infer them
    if size_col is None or quantity_col is None:
        # Look for rows that might contain size data
        for row in processed_table[header_row_idx + 1:]:
            if len(row) < 2:
                continue
                
            # Check if any cell contains a size identifier
            for i, cell in enumerate(row):
                if size_col is None and any(size in cell.upper() for size in ["XS", "S", "M", "L", "XL", "XXL"]):
                    size_col = i
                elif quantity_col is None and re.search(r'\d+', cell):
                    quantity_col = i
            
            # If we've found both columns, break
            if size_col is not None and quantity_col is not None:
                break
    
    # If we still couldn't determine the columns, return empty
    if size_col is None or quantity_col is None:
        return []
    
    # Extract the data rows
    size_data = []
    for row in processed_table[header_row_idx + 1:]:
        if len(row) <= max(size_col, quantity_col):
            continue
            
        size = row[size_col].strip()
        quantity_str = row[quantity_col].strip()
        
        # Skip empty rows
        if not size and not quantity_str:
            continue
            
        # Clean up the size
        size = re.sub(r'\s+', ' ', size)
        
        # Extract and clean the quantity
        quantity_match = re.search(r'([\d,]+(?:\.\d+)?)', quantity_str)
        if quantity_match:
            try:
                quantity = float(quantity_match.group(1).replace(',', ''))
                size_data.append({
                    'Size': size,
                    'Order Quantity': quantity
                })
            except ValueError:
                continue
    
    return size_data

def _find_header_row(table: List[List[str]]) -> int:
    """
    Find the header row in a table that contains size breakdown information.
    Returns the index of the header row, or -1 if not found.
    """
    for i, row in enumerate(table):
        if not row:
            continue
            
        row_text = " ".join(row).lower()
        if "size" in row_text and ("quantity" in row_text or "qty" in row_text):
            return i
    
    return -1

def _extract_size_breakdown_from_text(text: str) -> List[Dict[str, Any]]:
    """
    Extract size breakdown from text when table extraction fails.
    """
    size_data = []
    
    # Find the size breakdown section
    size_section_match = re.search(r'Size/Age Breakdown:(.*?)(?:ITL Factory Code:|Care Instruction Set|$)', text, re.DOTALL)
    if not size_section_match:
        return size_data
    
    size_text = size_section_match.group(1)
    
    # Try to extract size and quantity pairs
    # Pattern matches like "XS | XP | ECH | 165/64A 696"
    size_lines = re.findall(r'([A-Z]+(?:/[A-Z0-9]+)*)\s+(\d+)', size_text)
    
    for size, qty in size_lines:
        # Extract main size (first part before /)
        main_size = size.split('/')[0]
        
        try:
            quantity = float(qty)
            size_data.append({
                'Size': main_size,
                'Order Quantity': quantity
            })
        except ValueError:
            continue
    
    return size_data


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
    data['Color Code'] = extract_color_code(text)
    data['Size ID'] = extract_size_id(text)
    
    # NEW FIELDS: Delivery Date and Delivery Location
    data['Delivery Date'] = extract_delivery_date(text)
    data['Delivery Location'] = extract_delivery_location(text)
    
    # OTHER FIELDS: Address and Customer
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

def process_wo_file(wo_file):
    """Process WO file and extract all relevant data"""
    try:
        # Extract text from PDF
        text = extract_text_from_pdf(wo_file)
        if not text:
            st.error("Could not extract text from WO file")
            return None
        
        # Extract all relevant fields
        wo_data = {
            'po_number': extract_po_number(text),
            'color_code': extract_color_code(text),
            'factory_id': extract_factory_id(text),
            'date_of_mfr': extract_date_of_mfr(text),
            'vss_vsd': extract_vss_vsd(text),
            'silhouette': extract_silhouette(text),
            'product_code': extract_product_code(text),
            'care_instruction': extract_care_instruction(text),
            'season': clean_season_value(extract_season(text)),
            'quantity': extract_quantity(text),
            'delivery_date': extract_delivery_date(text),
            'size_id': extract_size_id(text),
            'delivery_location': extract_delivery_location(text),
            'customer': clean_customer_value(extract_customer(text)),
            'address': extract_address(text),
            'garment_components': extract_garment_components(text)
        }
        
        return wo_data
    except Exception as e:
        st.error(f"Error processing WO file: {str(e)}")
        return None
    

def extract_and_sort_wo_sizes(wo_items):
    """
    Extracts and sorts sizes from WO items table.
    Returns a list of dictionaries with size and quantity information.
    """
    if not wo_items:
        return []
    
    # Aggregate quantities by size
    size_quantities = {}
    for item in wo_items:
        size = (item.get('Size 1') or 
               item.get('Size') or 
               item.get('size') or '').strip()
        
        if not size:
            continue
            
        quantity = 0
        qty_str = (item.get('Quantity') or 
                  item.get('quantity') or '')
        
        if qty_str:
            try:
                quantity = float(str(qty_str).replace(',', ''))
            except (ValueError, TypeError):
                quantity = 0
        
        if size in size_quantities:
            size_quantities[size] += quantity
        else:
            size_quantities[size] = quantity
    
    # Convert to list of dictionaries
    size_list = []
    for size, quantity in size_quantities.items():
        size_list.append({
            'Size': size,
            'Order Quantity': int(quantity) if quantity == int(quantity) else quantity
        })
    
    # Sort by size
    size_order = {'XS': 0, 'S': 1, 'M': 2, 'L': 3, 'XL': 4, 'XXL': 5}
    size_list.sort(key=lambda x: size_order.get(x['Size'], 999))
    
    return size_list