import streamlit as st
from datetime import datetime
import os

def authenticate_user():
    """Handle user authentication and return selected user"""
    with st.sidebar:
        st.markdown("### ⭐ User Selection")
        selected_user = st.selectbox(
            "Select your name:",
            ["","👧udari", "👧Shaini", "👧Priyangi", "👧Uvini","👦Vihanga","👧Nimesha"],
            help="You must select a user to access the application"
        )
        
        if not selected_user:
            st.warning("⚠️ Please select a user to continue")
        
        # -------------------- PIN Authentication --------------------
        st.markdown("""
        <div class="pin-container">
            <div class="pin-title">🔐 Admin Access</div>
        </div>
        """, unsafe_allow_html=True)
        
        pin_input = st.text_input(
            "Enter PIN to access logs:",
            type="password",
            help="Enter the 4-digit PIN to access log search functionality"
        )
        
        # -------------------- Date Search (only visible if PIN is correct) --------------------
        if pin_input == "4391":
            st.markdown("""
            <div class="search-container">
                <div class="pin-title">🔍 Search Logs by Date</div>
            </div>
            """, unsafe_allow_html=True)
            
            search_date = st.date_input(
                "Select date to search:",
                help="Select the date for which you want to retrieve logs"
            )
            
            if st.button("Search Logs", key="search_logs"):
                from logging_utils import read_log_file_and_convert_to_excel
                date_str = search_date.strftime("%Y-%m-%d")
                excel_file, message = read_log_file_and_convert_to_excel(date_str)
                
                if excel_file:
                    st.success(message)
                    st.download_button(
                        label=f"⬇️ Download Log for {date_str}",
                        data=excel_file,
                        file_name=f"CS_AI_Tool_Log_{date_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.error(message)
        elif pin_input:
            st.error("❌ Incorrect PIN. Please try again.")
    
    return selected_user

def setup_sidebar():
    """Setup the sidebar with file uploaders"""
    selected_user = authenticate_user()
    
    with st.sidebar:
        st.markdown("""
        <div style="text-align: center; padding: 1rem 0; color: black;">
            <h2>⚙️ Control Panel</h2>
            <p style="opacity: 0.8;">Configure your analysis settings</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### 📁 File Upload")
        wo_file = st.file_uploader(
            "📄 Work Order (WO) PDF", 
            type="pdf",
            disabled=not selected_user,
            help="Upload your Work Order PDF file"
        )
        
        po_file = st.file_uploader(
            "📋 Purchase Order (PO) PDF", 
            type="pdf",
            disabled=not selected_user,
            help="Upload your Purchase Order PDF file"
        )
        if po_file:
            st.session_state.po_file = po_file
        if wo_file:
            st.success("✅ WO File Loaded")
        if po_file:
            st.success("✅ PO File Loaded")
        
        if wo_file and po_file:
            st.markdown("### 🚀 Ready to Process")
            st.info("Both files are loaded. Analysis will begin automatically.")
    
    return selected_user, wo_file, po_file