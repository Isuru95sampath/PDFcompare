import streamlit as st
import pdfplumber
import openpyxl
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO
from fuzzywuzzy import fuzz

# -------------------- UI Setup 
st.set_page_config(page_title=" PDF Merger CS", layout="centered")
st.title("📓 PDF Merger CS")

excel_file = st.file_uploader("📤 Upload Excel File", type=["xls", "xlsx"], key="excel")
pdf_file = st.file_uploader("📄 Upload PDF File", type=["pdf"], key="pdf")

# -------------------- Pattern 
style_pattern = re.compile(r"^\d{8}$")

# -------------------- Excel Reader 
def extract_styles_from_excel(file) -> list:
    try:
        if file.name.endswith(".xls"):
            df = pd.read_excel(file, sheet_name=0, header=None, engine="xlrd")
            for i in range(23, len(df)):
                val = df.iloc[i, 1]
                val_str = str(val).strip() if val is not None else ""
                if not val_str:
                    break
                if style_pattern.match(val_str):
                    return [val_str]  # ✅ Only return the first match
                else:
                    break
        else:
            wb = openpyxl.load_workbook(file, data_only=True)
            ws = wb.active
            row = 24
            while True:
                val = ws[f"B{row}"].value
                val_str = str(val).strip() if val is not None else ""
                if not val_str:
                    break
                if style_pattern.match(val_str):
                    return [val_str]
                else:
                    break
                row += 1
        return []
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return []

# -------------------- Style to PDF 
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

# -------------------- Merge PDF 
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

# -------------------- Main Logic 
if excel_file and pdf_file:
    styles = extract_styles_from_excel(excel_file)

    if styles:
        st.subheader("✅ Extracted Style Numbers")
        st.dataframe(pd.DataFrame(styles, columns=["Style"]))
        styles_pdf = create_styles_pdf(styles)
        final_pdf = merge_pdfs(pdf_file, styles_pdf)

        st.download_button(
            label="📄 Download Final PDF with Styles",
            data=final_pdf,
            file_name="Merged-PO.pdf",
            mime="application/pdf"
        )
    else:
        st.warning("No valid 8-digit style numbers found in the Excel file.")
elif excel_file or pdf_file:
    st.info("Please upload both an Excel and a PDF file to proceed.")
# -------------------- End of that code.................



def truncate_after_sri_lanka(addr: str) -> str:
    part, sep, _ = addr.partition("Sri Lanka")
    return (part + sep).strip() if sep else addr.strip()


def extract_wo_fields(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)


    customer = delivery = ""
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if "Deliver To:" in ln:
            customer = lines[i - 1].strip() if i > 0 else ""
            delivery = re.sub(r"Deliver To:\s*", "", ln).strip()
            break


    codes = []
    for line in lines:
        if "Product Code" in line:
            found = re.findall(r"Product Code[:\s]*([\w\s\-]+(?:\s*/\s*[\w\s\-]+)*)", line)
            for match in found:
                for code in match.split("/"):
                    clean = code.strip().upper()
                    if "-" in clean:
                        parts = clean.split("-")
                        if len(parts) >= 2:
                            clean = f"{parts[0]}-{parts[1]}"
                    if clean:
                        codes.append(clean)
            break


    po_numbers = list(set(re.findall(r'\b\d{7,8}\b', text)))
    return {
        "customer_name": customer,
        "delivery_address": delivery,
        "product_codes": list(set(codes)),
        "po_numbers": po_numbers
    }


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


    raw_addr = " ".join(address_lines)
    matches = re.findall(r".*Sri Lanka.*", text, re.IGNORECASE)
    unique = [raw_addr] + [m for m in matches if m != raw_addr]
    seen = []
    for a in unique:
        if a and a not in seen:
            seen.append(a)


    sri = [a for a in seen if "sri lanka" in a.lower()]
    chosen = max(sri, key=len) if sri else seen[0] if seen else raw_addr
    final_addr = truncate_after_sri_lanka(chosen)


    po_codes = re.findall(r"(LB\s*\d+)", text)


    sup_ref_codes = re.findall(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z]+[-\s]?\d+)", text, re.IGNORECASE)
    tag_codes = re.findall(r"TAG\.PRC\.TKT_(.*?)_REG", text)
    all_product_codes = list(set([c.strip().upper() for c in sup_ref_codes + tag_codes]))


    return {
        "delivery_location": final_addr,
        "product_codes": po_codes + all_product_codes,
        "all_found_addresses": seen
    }



# New helper function: Extract style numbers under "Extracted Style Numbers:" on first PO page
def extract_style_numbers_from_po_first_page(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        if len(pdf.pages) == 0:
            return []
        first_page = pdf.pages[0]
        text = first_page.extract_text() or ""
        lines = text.split("\n")


        extracted_styles = []
        capture = False
        for line in lines:
            if "Extracted Style Numbers:" in line:
                capture = True
                continue
            if capture:
                if line.strip() == "" or re.match(r"^\s*-+\s*$", line):
                    # Stop capturing on empty line or separator line
                    break
                # Extract style numbers: assume numbers are separated by spaces or commas
                numbers = re.findall(r"\b\d{6,8}\b", line)
                extracted_styles.extend(numbers)
        return extracted_styles
    

def reorder_wo_by_size(wo_items):
    size_order = {"XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5}

    def get_order(wo):
        size = wo.get("Size 1", "").strip().upper()
        return size_order.get(size, 99)  # Unknown sizes go last

    return sorted(wo_items, key=get_order)


def reorder_po_by_size(po_details):
    size_order = {"XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5}
   
    def get_order(po):
        size = po.get("Size", "").strip().upper()
        return size_order.get(size, 99)
   
    return sorted(po_details, key=get_order)


def extract_style_numbers_from_po_first_page(pdf_file):
    # Use PyMuPDF to read the first page for style numbers
    with fitz.open(stream=pdf_file.read(), filetype="pdf") as doc:
        first_page = doc[0].get_text()
        match = re.search(r"Extracted Style Numbers:\s*(.*)", first_page, re.IGNORECASE)
        if match:
            styles_text = match.group(1).strip()
            # Split by comma or newlines
            styles = re.split(r"[,\n]", styles_text)
            return [s.strip() for s in styles if s.strip()]
    return []

# -------------------- UPDATED PO DETAILS EXTRACTION FUNCTION --------------------
def extract_po_details(pdf_file):
    """
    Enhanced function to handle multiple PO formats:
    1. Original format with Sup Ref codes and item tables
    2. New format with TAG codes and Color/Size/Destination lines
    """
    pdf_file.seek(0)
    extracted_styles = extract_style_numbers_from_po_first_page(pdf_file)
    repeated_style = extracted_styles[0] if extracted_styles else ""

    pdf_file.seek(0)
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]

    # Try to detect which format this is
    has_tag_format = "TAG.PRC.TKT_" in text and "Color/Size/Destination :" in text
    has_original_format = any("Colour/Size/Destination:" in line for line in lines) or re.search(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z]+[-\s]?\d+)", text, re.IGNORECASE)

    po_items = []

    if has_tag_format and not has_original_format:
        # NEW FORMAT HANDLING - TAG format with Color/Size/Destination
        print("Detected TAG format PO")
        
        # Extract product code from TAG format
        tag_match = re.search(r"TAG\.PRC\.TKT_(.*?)_REG", text)
        product_code_used = tag_match.group(1).strip().upper() if tag_match else ""
        
        # Process each line for item extraction
        i = 0
        while i < len(lines):
            line = lines[i]
            
            # Look for item number pattern at start of line: "1 TAG.PRC.TKT_..."
            item_match = re.match(r'^(\d+)\s+TAG\.PRC\.TKT_.*?(\d+\.\d+)\s+PCS', line)
            if item_match:
                item_no = item_match.group(1)
                quantity = int(float(item_match.group(2)))
                
                # Look for Color/Size/Destination in next few lines
                colour = size = ""
                for j in range(i + 1, min(i + 5, len(lines))):
                    next_line = lines[j]
                    if "Color/Size/Destination :" in next_line:
                        # Extract color and size from format: "Color/Size/Destination : 34Y5 45K / L / X"
                        cs_part = next_line.split(":", 1)[1].strip()
                        cs_parts = [part.strip() for part in cs_part.split(" / ") if part.strip()]
                        
                        if len(cs_parts) >= 2:
                            # First part contains color code(s)
                            colour_part = cs_parts[0].strip()
                            # Extract the main color code (first part before space)
                            colour = colour_part.split()[0] if colour_part else ""
                            
                            # Second part is size
                            size = cs_parts[1].strip().upper()
                        break
                
                po_items.append({
                    "Item_Number": item_no,
                    "Item_Code": f"TAG_{product_code_used}",  # Create item code from TAG
                    "Quantity": quantity,
                    "Colour_Code": colour.upper() if colour else "",
                    "Size": size,
                    "Style 2": repeated_style,
                    "Product_Code": product_code_used,
                })
            i += 1

    else:
        # ORIGINAL FORMAT HANDLING (fallback to original logic)
        print("Detected original format PO or fallback")
        
        # Extract product codes (original logic)
        sup_ref_match = re.search(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z]+[-\s]?\d+)", text, re.IGNORECASE)
        sup_ref_code = sup_ref_match.group(1).strip().upper() if sup_ref_match else ""

        tag_code = ""
        for i, line in enumerate(lines):
            if "Item Description" in line:
                if i + 2 < len(lines):
                    second_line = lines[i + 2]
                    match = re.search(r"TAG\.PRC\.TKT_(.*?)_REG", second_line)
                    if match:
                        tag_code = match.group(1).strip().upper()
                break

        product_code_used = sup_ref_code if sup_ref_code else tag_code

        # Extract items (original logic)
        for i, line in enumerate(lines):
            item_match = re.match(r'^(\d+)\s+([A-Z0-9]+)\s+(\d+)\s+(\d+\.\d+)\s+PCS', line)
            if item_match:
                item_no, item_code, _, qty_str = item_match.groups()
                quantity = int(float(qty_str))
                colour = size = ""

                for j in range(i + 1, min(i + 10, len(lines))):
                    ln = lines[j]
                    if not colour and "Colour/Size/Destination:" in ln:
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

                po_items.append({
                    "Item_Number": item_no,
                    "Item_Code": item_code,
                    "Quantity": quantity,
                    "Colour_Code": (colour or "").strip().upper(),
                    "Size": (size or "").strip().upper(),
                    "Style 2": repeated_style,
                    "Product_Code": product_code_used,
                })

    return po_items


def extract_wo_items_table(pdf_file, product_codes=None):
    import re
    items = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                for row in table or []:
                    if row and len(row) >= 6:
                        style = (row[0] or "").strip()
                        if re.match(r"^\d{8}$", style):
                            colour = (row[1] or "").strip().upper()
                            qty = 0
                            size_val = ""


                            for col in row[2:]:
                                if col:
                                    text = str(col).strip().upper()
                                    if "|" in text:
                                        left_size = text.split("|")[0].strip()
                                        if left_size:
                                            size_val = left_size
                                            break
                                    if "/" in text and not size_val:
                                        left_size = text.split("/")[0].strip()
                                        if left_size:
                                            size_val = left_size
                                            break
                                    if not size_val and not text.isdigit():
                                        size_val = text
                                        break


                            for col in reversed(row):
                                if col and str(col).strip().isdigit():
                                    qty = int(str(col).strip())
                                    break


                            if qty > 0:
                                items.append({
                                    "Style": style,
                                    "WO Colour Code": colour,
                                    "Size 1": size_val,
                                    "Quantity": qty,
                                    "WO Product Code": " / ".join(product_codes) if product_codes else ""
                                })
    return items


def enhanced_quantity_matching(wo_items, po_details, tolerance=0):
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

            # Full Match Check
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
                # Score partial match
                score = 0
                if abs(pq - wq) <= tolerance:
                    score += 1
                if ws == ps:
                    score += 1
                if wc == pc:
                    score += 1
                if wstyle == pstyle:
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

            qty_match = "👍 Yes" if abs(pq - wq) <= tolerance else "❌ No"
            size_match = "👍 Yes" if ws == ps else "❌ No"
            colour_match = "👍 Yes" if wc == pc else "❌ No"
            style_match = "👍 Yes" if wstyle == pstyle else "❌ No"

            matched.append({
                "Style": wstyle, "Style 2": pstyle,
                "WO Size": ws, "PO Size": ps,
                "WO Colour Code": wc, "PO Colour Code": pc,
                "WO Qty": wq, "PO Qty": pq,
                "Qty Match": qty_match, "Size Match": size_match,
                "Colour Match": colour_match, "Style Match": style_match,
                "Diff": pq - wq,
                "Status": "🟨 NO Match",
                "PO Item Code": po.get("Item_Code", "")
            })
            used.add(partial_match_idx)
        else:
            # No PO match found for this WO item
            mismatched.append({
                "Style": wstyle, "Style 2": "",
                "WO Size": ws, "PO Size": "",
                "WO Colour Code": wc, "PO Colour Code": "",
                "WO Qty": wq, "PO Qty": None,
                "Qty Match": "❌ No", "Size Match": "❌ No",
                "Colour Match": "❌ No", "Style Match": "❌ No",
                "Diff": "", "Status": "❌ No PO Match", "PO Item Code": ""
            })

    # Remaining unmatched PO items
    for idx, po in enumerate(po_details):
        if idx not in used:
            ps = po.get("Size", "").strip().upper()
            pc = po.get("Colour_Code", "").strip().upper()
            pstyle = po.get("Style 2", "").strip()
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

    return matched, mismatched


def compare_addresses(wo, po):
    ns = fuzz.token_sort_ratio(wo["customer_name"], po["delivery_location"])
    as_ = fuzz.token_sort_ratio(wo["delivery_address"], po["delivery_location"])
    comb = max(ns, as_)
    return {"WO Name": wo["customer_name"], "WO Addr": wo["delivery_address"], "PO Addr": po["delivery_location"],
            "Name %": ns, "Addr %": as_, "Overall %": comb, "Status": "✅ Match" if comb > 85 else "⚠️ Review"}


def compare_codes(po_details, wo_items):
    po_codes = set(po.get("Product_Code", "").strip().upper() for po in po_details if po.get("Product_Code"))
    wo_codes = set(w.get("WO Product Code", "").strip().upper() for w in wo_items if w.get("WO Product Code"))


    comparison = []
    all_codes = po_codes.union(wo_codes)


    for code in all_codes:
        in_po = code in po_codes
        in_wo = code in wo_codes
        status = "✅Match" if in_po and in_wo else "❌ Missing in WO" if in_po else "❌ Missing in PO"
        comparison.append({"PO Code": code if in_po else "", "WO Code": code if in_wo else "", "Status": status})


    return comparison


# --- Streamlit UI ---
st.set_page_config(page_title="WO ↔ PO Comparator", layout="wide")


st.title("📄 Customer Care System")
st.subheader("🔁 PO vs WO Comparison Dashboard")


with st.sidebar:
    st.header("⚙️ Settings")
    method = st.selectbox("Select Matching Method:",
                          ["Enhanced Matching (with PO Color/Size)", "Smart Matching (Exact)", "Smart Matching with Tolerance"])
    wo_file = st.file_uploader("📤 Upload WO PDF", type="pdf")
    po_file = st.file_uploader("📤 Upload PO PDF", type="pdf")


if wo_file and po_file:
    with st.spinner("🔄 Processing files..."):
        wo = extract_wo_fields(wo_file)
        po = extract_po_fields(po_file)
        wo_items = extract_wo_items_table(wo_file, wo["product_codes"])
        wo_items = reorder_wo_by_size(wo_items)  # <-- reorder WO items here

        po_details_raw = extract_po_details(po_file)
        po_details = reorder_po_by_size(po_details_raw)
        addr_res = compare_addresses(wo, po)
        code_res = compare_codes(po_details, wo_items)

        if "Enhanced" in method:
            matched, mismatched = enhanced_quantity_matching(wo_items, po_details)
        else:
            matched, mismatched = [], []

    st.success("✅ Comparison Completed")

    st.markdown("---")
    st.subheader("🧭 Address Comparison")
    st.dataframe(pd.DataFrame([addr_res]), use_container_width=True)

    st.subheader("🔢 Product Code Comparison")

    po_all_codes = [po.get("Product_Code", "").strip().upper() for po in po_details if po.get("Product_Code")]
    wo_all_codes = [wo.get("WO Product Code", "").strip().upper() for wo in wo_items if wo.get("WO Product Code")]

    comparison_rows = []
    max_len = max(len(po_all_codes), len(wo_all_codes)) if po_all_codes or wo_all_codes else 0

    for i in range(max_len):
        po_code = po_all_codes[i] if i < len(po_all_codes) else ""
        wo_code = wo_all_codes[i] if i < len(wo_all_codes) else ""

        if po_code and wo_code and po_code == wo_code:
            status = "✅ Exact Match"
        elif po_code and wo_code and "/" in wo_code:
            wo_parts = [part.strip().upper() for part in wo_code.split("/")]
            status = "✅ Exact Match" if po_code in wo_parts else "❌ No Match"
        elif po_code and wo_code and "/" in po_code:
            po_parts = [part.strip().upper() for part in po_code.split("/")]
            status = "✅ Partial Match (PO contains WO code)" if wo_code in po_parts else "❌ No Match"
        elif po_code and wo_code:
            status = "❌ No Match"
        else:
            status = "⚪ Empty"

        comparison_rows.append({
            "PO Product Code": po_code,
            "WO Product Code": wo_code,
            "Match Status": status
        })

    code_table_df = pd.DataFrame(comparison_rows)
    st.dataframe(code_table_df, use_container_width=True)

    st.subheader("🌇 Matched WO/PO Items")
    if matched:
        st.dataframe(pd.DataFrame(matched), use_container_width=True)
    else:
        st.info("No matched items found or matching method not selected.")

    # --- New logic: show alert if all checks are good ---

    address_ok = addr_res.get("Status", "") == "✅ Match"
    codes_ok = (not code_table_df.empty) and all(code_table_df["Match Status"].str.contains("✅"))
    matched_df = pd.DataFrame(matched) if matched else pd.DataFrame()
    matched_ok = (not matched_df.empty) and all(matched_df["Status"].str.contains("Full Match"))
    mismatched_empty = len(mismatched) == 0

    if address_ok and codes_ok and matched_ok and mismatched_empty:
        st.success("🎉 All checking data are matched successfully!")
    else:
        st.warning("⚠️ Some mismatches detected. Please review the details above.")

    
    # Safe definitions to avoid NameError
    matched_df = pd.DataFrame(matched) if matched else pd.DataFrame()
    code_table_df = pd.DataFrame(comparison_rows) if comparison_rows else pd.DataFrame()
    address_ok = addr_res.get("Status", "") == "✅ Match"

    # Checks
    codes_ok = not code_table_df.empty and all(code_table_df["Match Status"] == "✅ Exact Match")
    matched_ok = not matched_df.empty and all(matched_df["Status"] == "🟩 Full Match")

    if address_ok and codes_ok and matched_ok:
        st.success("🎉 All data in Address, Product Codes, and Matched Items are fully matched!")

    st.subheader("❗ Mismatched or Extra Items")
    if mismatched:
        st.dataframe(pd.DataFrame(mismatched), use_container_width=True)
    else:
        st.success("No mismatched items found.")


    st.subheader("🧾 Work Order (WO) Items Table")
    st.dataframe(pd.DataFrame(wo_items), use_container_width=True)


    st.subheader("📦 Purchase Order (PO) Details")
    st.dataframe(pd.DataFrame(po_details), use_container_width=True)


st.markdown("<br><hr><center><b style='color:#888'>Created by Razz... </b></center>",
            unsafe_allow_html=True)