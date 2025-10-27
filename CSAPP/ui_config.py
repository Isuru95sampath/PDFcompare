import streamlit as st

def configure_page():
    """Configure the Streamlit page settings"""
    st.set_page_config(
        page_title="Bandix Data Entry Checking Tool – Price Tickets - Razz Solutions",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )

def apply_custom_css():
    """Apply custom CSS styling to the Streamlit app"""
    st.markdown("""
    <style>
        /* Import Google Fonts */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        
        /* Global Styles */
        .main {
            font-family: 'Inter', sans-serif;
        }
        
        /* Header Styling */
        .main-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 2rem;
            border-radius: 15px;
            margin-bottom: 2rem;
            text-align: center;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        }
        
        .main-title {
            color: white;
            font-size: 2.5rem;
            font-weight: 700;
            margin-bottom: 0.5rem;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        
        .main-subtitle {
            color: rgba(255,255,255,0.9);
            font-size: 1.2rem;
            font-weight: 400;
            margin: 0;
        }
        
        /* Status Cards */
        .status-card {
            background: white;
            padding: 1.5rem;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.08);
            border-left: 4px solid #4CAF50;
            margin: 1rem 0;
        }
        
        .status-card.warning {
            border-left-color: #FF9800;
        }
        
        .status-card.error {
            border-left-color: #f44336;
        }
        
        /* Section Headers */
        .section-header {
            background: linear-gradient(90deg, #f8f9fa 0%, #e9ecef 100%);
            padding: 1rem 1.5rem;
            border-radius: 10px;
            border-left: 5px solid #007bff;
            margin: 1.5rem 0 1rem 0;
        }
        
        .section-title {
            color: #2c3e50;
            font-size: 1.3rem;
            font-weight: 600;
            margin: 0;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }
        
        /* Upload Area Styling */
        .upload-container {
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            border: 2px dashed #007bff;
            border-radius: 15px;
            padding: 2rem;
            text-align: center;
            margin: 1rem 0;
            transition: all 0.3s ease;
        }
        
        .upload-container:hover {
            border-color: #0056b3;
            background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
        }
        
        /* Metrics Cards */
        .metric-card {
            background: white;
            padding: 1.5rem;
            border-radius: 12px;
            text-align: center;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
            border-top: 4px solid #007bff;
            transition: transform 0.3s ease;
        }
        
        .metric-card:hover {
            transform: translateY(-5px);
        }
        
        .metric-value {
            font-size: 2rem;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 0.5rem;
        }
        
        .metric-label {
            color: #6c757d;
            font-weight: 500;
            text-transform: uppercase;
            font-size: 0.9rem;
            letter-spacing: 1px;
        }
        
        /* Success/Warning/Error Messages */
        .alert-success {
            background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
            border: 1px solid #c3e6cb;
            color: #155724;
            padding: 1rem 1.5rem;
            border-radius: 10px;
            margin: 1rem 0;
            font-weight: 500;
        }
        
        .alert-warning {
            background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%);
            border: 1px solid #ffeaa7;
            color: #856404;
            padding: 1rem 1.5rem;
            border-radius: 10px;
            margin: 1rem 0;
            font-weight: 500;
        }
        
        .alert-info {
            background: linear-gradient(135deg, #d1ecf1 0%, #bee5eb 100%);
            border: 1px solid #bee5eb;
            color: #0c5460;
            padding: 1rem 1.5rem;
            border-radius: 10px;
            margin: 1rem 0;
            font-weight: 500;
        }
        
        /* Table Styling */
        .dataframe {
            border: none !important;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1) !important;
            border-radius: 10px !important;
            overflow: hidden !important;
        }
        
        /* Sidebar Styling */
        .css-1d391kg {
            background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
        }
        
        .css-1d391kg .css-1v0mbdj {
            color: white;
        }
        
        /* Footer */
        .footer {
            text-align: center;
            padding: 2rem;
            margin-top: 3rem;
            border-top: 2px solid #e9ecef;
            color: #6c757d;
            font-style: italic;
        }
        
        /* Loading Spinner */
        .loading-container {
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 2rem;
        }
        
        /* Progress Steps */
        .progress-steps {
            display: flex;
            justify-content: center;
            margin: 2rem 0;
            gap: 1rem;
        }
        
        .step {
            padding: 0.5rem 1rem;
            border-radius: 20px;
            background: #e9ecef;
            color: #6c757d;
            font-weight: 500;
            font-size: 0.9rem;
        }
        
        .step.active {
            background: #007bff;
            color: white;
        }
        
        .step.completed {
            background: #28a745;
            color: white;
        }
        
        /* Mismatch Summary Card */
        .mismatch-summary {
            background: linear-gradient(135deg, #fff5f5 0%, #ffe0e0 100%);
            border: 1px solid #ffcdd2;
            color: #c62828;
            padding: 1rem 1.5rem;
            border-radius: 10px;
            margin: 1rem 0;
            font-weight: 500;
        }
        
        .mismatch-example {
            background: white;
            padding: 1rem;
            border-radius: 8px;
            margin-top: 1rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        /* PIN Input Styling */
        .pin-container {
            background: rgba(255, 255, 255, 0.1);
            border-radius: 10px;
            padding: 1rem;
            margin-top: 1rem;
            border: 1px solid rgba(255, 255, 255, 0.2);
        }
        
        .pin-title {
            color: black;  /* Changed from white to black */
            font-size: 1rem;
            font-weight: 600;
            margin-bottom: 0.5rem;
        }
        
        .search-container {
            background: rgba(255, 255, 255, 0.1);
            border-radius: 10px;
            padding: 1rem;
            margin-top: 1rem;
            border: 1px solid rgba(255, 255, 255, 0.2);
        }
    </style>
    """, unsafe_allow_html=True)

def display_header():
    """Display the main header of the application"""
    st.markdown("""
    <div class="main-header">
        <h1 class="main-title">🚀 Brandix Data Entry Checking Tool – Price Tickets</h1>
        <p class="main-subtitle">Advanced PO vs WO Comparison Dashboard | Powered by Razz </p>
    </div>
    """, unsafe_allow_html=True)

def display_footer():
    """Display the footer of the application"""
    st.markdown("""
    <div class="footer">
        <p>
            🚀 <strong>Customer Care System v2.0</strong> | 
            Powered by <strong>Razz....</strong> | 
            Advanced PDF Analysis & Comparison Technology
        </p>
    </div>
    """, unsafe_allow_html=True)