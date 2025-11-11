import streamlit as st
import pandas as pd
from datetime import datetime

def initialize_page():
    """Initialize page configuration"""
    st.set_page_config(page_title="WO & PO Comparison System for LB 5801", layout="wide")

def initialize_session_state():
    """Initialize session state variables"""
    if 'checker_name' not in st.session_state:
        st.session_state.checker_name = None
    if 'wo_data' not in st.session_state:
        st.session_state.wo_data = None
    if 'po_data' not in st.session_state:
        st.session_state.po_data = None

def create_sidebar():
    """Create sidebar with navigation"""
    st.sidebar.title("Navigation")

    # Checker name selection (must be selected first)
    checker_names = [
        "Select Checker",
        "Vihange Perera",
        "Tarini Alwis",
        "Uvini Perera",
        "Udari Liyanage",
        "Venura Prabashwara",
        "Shaini Nuwandhara",
        "Nimesha Samarathunga",
        "Priyangika Damayanthi"
    ]

    selected_checker = st.sidebar.selectbox("Select Checker Name", checker_names)
    if selected_checker != "Select Checker":
        st.session_state.checker_name = selected_checker
    else:
        st.session_state.checker_name = None

    st.sidebar.markdown("---")

    if st.session_state.checker_name is None:
        st.sidebar.warning("⚠️ Please select a checker name first to continue")

    # Radio buttons for navigation
    page = st.sidebar.radio(
        "Select Page",
        ["Merge PO", "WO & PO Analysis"],  # Changed to have only 2 pages
        disabled=(st.session_state.checker_name is None)
    )
    
    return page

def display_wo_details(wo_data, checker_name):
    """Display WO details in the UI"""
    st.success("✅ WO processed successfully!")

    # Display WO data
    st.subheader("📋 WO Details")
    
    # Create a dataframe from the WO data
    df_data = []
    
    for key, value in wo_data.items():
        if key not in ['Size Breakdown', 'Size Breakdown Table']:
            df_data.append({
                'Field': key,
                'Value': value
            })
    
    df = pd.DataFrame(df_data)
    st.dataframe(df, use_container_width=True)