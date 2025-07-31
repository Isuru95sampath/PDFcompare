import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from fuzzywuzzy import fuzz

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

    # ✅ Extract all product codes under "Product Code:"
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

    # Also extract product codes for Product Code Comparison display
    sup_ref_codes = re.findall(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z]+[-\s]?\d+)", text, re.IGNORECASE)
    tag_codes = re.findall(r"TAG\.PRC\.TKT_(.*?)_REG", text)
    all_product_codes = list(set([c.strip().upper() for c in sup_ref_codes + tag_codes]))

    return {
        "delivery_location": final_addr,
        "product_codes": po_codes + all_product_codes,
        "all_found_addresses": seen
    }

def reorder_po_by_size(po_details):
    size_order = {"XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5}
    
    def get_order(po):
        size = po.get("Size", "").strip().upper()
        return size_order.get(size, 99)  # Unknown sizes go to the bottom
    
    return sorted(po_details, key=get_order)


def extract_po_details(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    # Method 1: Sup. Ref.
    sup_ref_match = re.search(r"Sup\.?\s*Ref\.?\s*[:\-]?\s*([A-Z]+[-\s]?\d+)", text, re.IGNORECASE)
    sup_ref_code = sup_ref_match.group(1).strip().upper() if sup_ref_match else ""

    # Method 2: Find TAG.PRC.TKT_*REG pattern under Item Description
    tag_code = ""
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    for i, line in enumerate(lines):
        if "Item Description" in line:
            if i + 2 < len(lines):
                second_line = lines[i + 2]
                match = re.search(r"TAG\.PRC\.TKT_(.*?)_REG", second_line)
                if match:
                    tag_code = match.group(1).strip().upper()
            break

    product_code_used = sup_ref_code if sup_ref_code else tag_code
    product_code_note = "(From Sup. Ref.)" if sup_ref_code else "(From TAG.PRC.TKT_)"

    po_items = []
    for i, line in enumerate(lines):
        item_match = re.match(r'^(\d+)\s+([A-Z0-9]+)\s+(\d+)\s+(\d+\.\d+)\s+PCS', line)
        if item_match:
            item_no, item_code, _, qty_str = item_match.groups()
            quantity = int(float(qty_str))
            colour = size = ""
            style_2 = ""

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

                if not style_2:
                    match = re.search(r"(\d{6,})\s*$", ln)
                    if match:
                        style_2 = match.group(1)

            po_items.append({
                "Item_Number": item_no,
                "Item_Code": item_code,
                "Quantity": quantity,
                "Colour_Code": (colour or "").strip().upper(),
                "Size": (size or "").strip().upper(),
                "Style 2": style_2,
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
                        if re.match(r"^\d{8}$", style):  # Valid style
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

        found_exact = False
        for idx, po in enumerate(po_details):
            if idx in used:
                continue
            pq = po["Quantity"]
            ps = po.get("Size", "").strip().upper()
            pc = po.get("Colour_Code", "").strip().upper()
            pstyle = po.get("Style 2", "").strip()

            if wq == pq and ws == ps and wc == pc and wstyle == pstyle:
                matched.append({
                    **{"Style": wstyle, "WO Size": ws, "PO Size": ps, "WO Colour Code": wc, "PO Colour Code": pc,
                       "WO Qty": wq, "PO Qty": pq, "Style 2": pstyle},
                    **{"Qty Match": "Yes", "Size Match": "Yes", "Colour Match": "Yes", "Style Match": "Yes",
                       "Diff": 0, "Status": "✅ Full Match", "PO Item Code": po.get("Item_Code", "")}
                })
                used.add(idx)
                found_exact = True
                break

        if found_exact:
            continue

        best_idx, best_diff = None, float("inf")
        for idx, po in enumerate(po_details):
            if idx in used:
                continue
            diff = abs(po["Quantity"] - wq)
            if diff <= tolerance and diff < best_diff:
                best_idx = idx
                best_diff = diff

        if best_idx is not None:
            po = po_details[best_idx]
            used.add(best_idx)
            pq = po["Quantity"]
            ps = po.get("Size", "").strip().upper()
            pc = po.get("Colour_Code", "").strip().upper()
            pstyle = po.get("Style 2", "").strip()

            qty_match = "Yes" if wq == pq else "No"
            size_match = "Yes" if ws == ps else "No"
            colour_match = "Yes" if wc == pc else "No"
            style_match = "Yes" if wstyle == pstyle else "No"

            full = all([qty_match == "Yes", size_match == "Yes", colour_match == "Yes", style_match == "Yes"])

            matched.append({
                **{"Style": wstyle, "WO Size": ws, "PO Size": ps, "WO Colour Code": wc, "PO Colour Code": pc,
                   "WO Qty": wq, "PO Qty": pq, "Style 2": pstyle},
                **{"Qty Match": qty_match, "Size Match": size_match, "Colour Match": colour_match, "Style Match": style_match,
                   "Diff": pq - wq,
                   "Status": "✅ Full Match" if full else "❌ Partial Match",
                   "PO Item Code": po.get("Item_Code", "")}
            })
        else:
            mismatched.append({
                **{"Style": wstyle, "WO Size": ws, "PO Size": "", "WO Colour Code": wc, "PO Colour Code": "",
                   "WO Qty": wq, "PO Qty": None, "Style 2": ""},
                **{"Qty Match": "No", "Size Match": "No", "Colour Match": "No", "Style Match": "No",
                   "Diff": "", "Status": "❌ No PO Qty", "PO Item Code": ""}
            })

    for idx, po in enumerate(po_details):
        if idx not in used:
            ps = po.get("Size", "").strip().upper()
            pc = po.get("Colour_Code", "").strip().upper()
            pstyle = po.get("Style 2", "").strip()
            mismatched.append({
                **{"Style": "N/A", "WO Size": "N/A", "PO Size": ps, "WO Colour Code": "N/A", "PO Colour Code": pc,
                   "WO Qty": None, "PO Qty": po["Quantity"], "Style 2": pstyle},
                **{"Qty Match": "No", "Size Match": "No", "Colour Match": "No", "Style Match": "No",
                   "Diff": "", "Status": "❌ Extra PO Item", "PO Item Code": po.get("Item_Code", "")}
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
        status = "✅ Match" if in_po and in_wo else "❌ Missing in WO" if in_po else "❌ Missing in PO"
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
        po_details = extract_po_details(po_file)
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
    st.subheader("📍 Address Comparison")
    st.dataframe(pd.DataFrame([addr_res]), use_container_width=True)

    st.subheader("🔢 Product Code Comparison")

    # Extract ALL product codes from PO Details table (including duplicates)
    po_all_codes = []
    for po in po_details:
        code = po.get("Product_Code", "").strip().upper()
        if code:
            po_all_codes.append(code)
    
    # Extract ALL product codes from WO Items table (including duplicates)
    wo_all_codes = []
    for wo in wo_items:
        code = wo.get("WO Product Code", "").strip().upper()
        if code:
            wo_all_codes.append(code)

    # Create detailed comparison table with match status
    comparison_rows = []
    max_len = max(len(po_all_codes), len(wo_all_codes)) if po_all_codes or wo_all_codes else 0
    
    for i in range(max_len):
        po_code = po_all_codes[i] if i < len(po_all_codes) else ""
        wo_code = wo_all_codes[i] if i < len(wo_all_codes) else ""
        
        # Check for exact match
        if po_code and wo_code and po_code == wo_code:
            status = "✅ Exact Match"
        # Check if WO code contains "/" and any part matches PO code
        elif po_code and wo_code and "/" in wo_code:
            wo_parts = [part.strip().upper() for part in wo_code.split("/")]
            if po_code in wo_parts:
                status = "✅ Exact Match"
            else:
                status = "❌ No Match"
        # Check if PO code contains "/" and any part matches WO code  
        elif po_code and wo_code and "/" in po_code:
            po_parts = [part.strip().upper() for part in po_code.split("/")]
            if wo_code in po_parts:
                status = "✅ Partial Match (PO contains WO code)"
            else:
                status = "❌ No Match"
        # Both codes exist but no match
        elif po_code and wo_code:
            status = "❌ No Match"
        # One or both codes are empty
        else:
            status = "⚪ Empty"
        
        comparison_rows.append({
            "PO Product Code": po_code,
            "WO Product Code": wo_code,
            "Match Status": status
        })

    # Create comparison table
    code_table_df = pd.DataFrame(comparison_rows)

    # Show the table
    st.dataframe(code_table_df, use_container_width=True)

    # Also show the original comparison logic
   

    st.subheader("✅ Matched WO/PO Items")
    if matched:
        st.dataframe(pd.DataFrame(matched), use_container_width=True)
    else:
        st.info("No matched items found or matching method not selected.")

    address_ok = addr_res.get("Overall %", 0) == 100
    codes_df = pd.DataFrame(code_res)
    codes_ok = not codes_df.empty and all(codes_df["Status"] == "✅ Match")
    matched_df = pd.DataFrame(matched)
    status_ok = not matched_df.empty and all(matched_df["Status"] == "✅ Full Match")

    if address_ok and codes_ok and status_ok:
        st.success("🎉 All items are FULLY MATCHED between PO and WO!")

    st.subheader("❗ Mismatched or Extra Items")
    if mismatched:
        st.dataframe(pd.DataFrame(mismatched), use_container_width=True)
    else:
        st.success("No mismatched or extra PO/WO items found.")

    st.subheader("🧾 Work Order (WO) Items Table")
    st.dataframe(pd.DataFrame(wo_items), use_container_width=True)

    st.subheader("📦 Purchase Order (PO) Details")
    st.dataframe(pd.DataFrame(po_details), use_container_width=True)

st.markdown("<br><hr><center><b style='color:#888'>Created by Razz... </b></center>",
            unsafe_allow_html=True)