import re
import os
import tempfile
import streamlit as st
from collections import OrderedDict
import time

# Lazy loading of heavy dependencies
@st.cache_resource
def load_dependencies():
    """Load heavy dependencies only when needed"""
    deps = {}
    try:
        import PyPDF2
        deps['PyPDF2'] = PyPDF2
    except ModuleNotFoundError:
        st.error("Module 'PyPDF2' not found. Install with: pip install PyPDF2")
        st.stop()
    
    try:
        import extract_msg
        deps['extract_msg'] = extract_msg
    except ModuleNotFoundError:
        st.error("Module 'extract_msg' not found. Install with: pip install extract-msg")
        st.stop()
    
    try:
        from xhtml2pdf import pisa
        deps['pisa'] = pisa
    except ModuleNotFoundError:
        st.error("Module 'xhtml2pdf' not found. Install with: pip install xhtml2pdf")
        st.stop()
    
    try:
        from PyPDF2 import PdfMerger, PdfReader
        deps['PdfMerger'] = PdfMerger
        deps['PdfReader'] = PdfReader
    except ModuleNotFoundError:
        st.error("Module 'PyPDF2' not found. Install with: pip install PyPDF2")
        st.stop()
    
    try:
        import email
        deps['email'] = email
    except ModuleNotFoundError:
        st.error("Module 'email' not found.")
        st.stop()
    
    try:
        from bs4 import BeautifulSoup
        deps['BeautifulSoup'] = BeautifulSoup
    except ModuleNotFoundError:
        st.error("Module 'beautifulsoup4' not found. Install with: pip install beautifulsoup4")
        st.stop()
    
    return deps

def extract_email_details(file_path, deps):
    subject, factory_code, coo = "", "", ""
    attachments = []
    html_content = ""

    if file_path.lower().endswith(".msg"):
        try:
            msg = deps['extract_msg'].Message(file_path)
            subject = msg.subject or "No Subject"
            html_content = msg.htmlBody or msg.body or ""
            if isinstance(html_content, bytes):
                html_content = html_content.decode(errors="ignore")
            for att in msg.attachments:
                if att.longFilename and att.longFilename.lower().endswith(".pdf"):
                    att_path = os.path.join(tempfile.gettempdir(), att.longFilename)
                    if not os.path.exists(att_path):
                        with open(att_path, "wb") as f:
                            f.write(att.data)
                    attachments.append(att_path)
        except Exception as e:
            st.error(f"Error reading .msg file: {e}")

    elif file_path.lower().endswith(".eml"):
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                msg = deps['email'].message_from_file(f)
            subject = msg.get("subject", "No Subject")

            for part in msg.walk():
                ctype = part.get_content_type()
                disp = str(part.get("Content-Disposition") or "")

                if ctype == "text/html" and not html_content:
                    payload = part.get_payload(decode=True)
                    html_content = payload.decode(errors="ignore") if isinstance(payload, bytes) else str(payload)
                elif ctype == "text/plain" and not html_content:
                    payload = part.get_payload(decode=True)
                    html_content = payload.decode(errors="ignore") if isinstance(payload, bytes) else str(payload)
                elif "attachment" in disp and ctype == "application/pdf":
                    filename = part.get_filename()
                    if filename:
                        att_path = os.path.join(tempfile.gettempdir(), filename)
                        if not os.path.exists(att_path):
                            with open(att_path, "wb") as f:
                                f.write(part.get_payload(decode=True))
                        attachments.append(att_path)
        except Exception as e:
            st.error(f"Error reading .eml file: {e}")

    # Extract Factory Code & COO
    text_content = deps['BeautifulSoup'](html_content, "html.parser").get_text(" ", strip=True)
    fc_match = re.search(r"(Factory\s*Code[:\-]?\s*)([A-Za-z0-9\-_\/]+)", text_content, re.IGNORECASE)
    coo_match = re.search(r"(COO|Country\s*of\s*Origin)[:\-]?\s*([A-Za-z ]+)", text_content, re.IGNORECASE)

    if fc_match:
        factory_code = fc_match.group(2).strip()
    if coo_match:
        coo = coo_match.group(2).strip()

    return subject, factory_code, coo, html_content, attachments

def convert_html_to_pdf(source_html, output_filename, pisa):
    if not source_html:
        source_html = "<p>No content available.</p>"
    if isinstance(source_html, bytes):
        source_html = source_html.decode(errors="ignore")

    with open(output_filename, "wb") as output_file:
        pisa_status = pisa.CreatePDF(source_html, dest=output_file, encoding="utf-8")
    return pisa_status.err

def create_pdf_from_html(html_content, output_path, subject="", factory_code="", coo="", pisa=None):
    header_html = f"""
    <div style="font-family: Arial, sans-serif; padding:10px;">
    <h2 style="color:#2c3e50;">📧 Email Details</h2>
    <p><b>Subject:</b> {subject}</p>
    """
    if factory_code:
        header_html += f"<p><b>Factory Code:</b> {factory_code}</p>"
    if coo:
        header_html += f"<p><b>COO:</b> {coo}</p>"

    header_html += "<hr><h3>Email Body:</h3>"

    body_html = html_content or "<p>No content available.</p>"

    table_style = """
    <style>
    body, table, td, th { font-family: Arial; font-size: 11pt; color: #000; }
    table { border-collapse: collapse; width: 100%; margin: 8px 0; }
    th, td { border: 1px solid #000; padding: 6px; text-align: left; vertical-align: top; word-wrap: break-word; }
    th { background-color: #f0f0f0; font-weight: bold; }
    tr:nth-child(even) { background-color: #fafafa; }
    </style>
    """

    final_html = table_style + header_html + body_html + "</div>"
    convert_html_to_pdf(final_html, output_path, pisa)

def merge_pdfs(pdf_list, output_path, PdfMerger, PdfReader):
    merger = PdfMerger()
    for pdf in pdf_list:
        try:
            with open(pdf, "rb") as f:
                merger.append(PdfReader(f, strict=False))
        except Exception as e:
            st.warning(f"⚠️ Skipping invalid PDF {pdf}: {e}")
    merger.write(output_path)
    merger.close()

def process_email_to_pdf(email_file):
    """Process email file and return merged PDF"""
    # Create a dedicated temporary directory for this operation to avoid permission errors
    temp_dir = tempfile.mkdtemp()
    st.info(f"🔧 Using temporary directory: `{temp_dir}`")

    # Helper function to create a safe filename
    def sanitize_filename(filename):
        return re.sub(r'[\\/*?:"<>|]', "", filename)

    try:
        with st.spinner("⏳ Processing email..."):
            # Load dependencies only when needed
            deps = load_dependencies()
            
            # Sanitize the uploaded filename to remove problematic characters
            safe_filename = sanitize_filename(email_file.name)
            email_path = os.path.join(temp_dir, safe_filename)
            
            st.write(f"🔧 Attempting to save email to: `{email_path}`")

            # Try to save the uploaded file to our safe temp directory
            try:
                with open(email_path, "wb") as f:
                    f.write(email_file.read())
            except PermissionError as e:
                st.error(f"❌ CRITICAL: Permission denied when trying to save the email file.")
                st.error(f"Your system's security software (Antivirus) is likely blocking access to the temp folder.")
                st.error(f"Please add an exclusion for Python or the temp folder in your Antivirus settings.")
                st.exception(e)
                return None

            subject, factory_code, coo, html_content, attachments = extract_email_details(email_path, deps)

            if not attachments:
                st.error("⚠️ No PO PDF attachments found in the email.")
                return None

            email_pdf = os.path.join(temp_dir, f"{safe_filename}_content.pdf")
            create_pdf_from_html(html_content, email_pdf, subject, factory_code, coo, deps['pisa'])

            merged_pdf = os.path.join(temp_dir, f"Merged_PO_{factory_code or 'Unknown'}_{safe_filename}.pdf")
            merge_pdfs([email_pdf] + attachments, merged_pdf, deps['PdfMerger'], deps['PdfReader'])

            return merged_pdf, temp_dir
    
    except Exception as e:
        st.error(f"❌ An unexpected error occurred during processing.")
        st.exception(e)
        return None, temp_dir