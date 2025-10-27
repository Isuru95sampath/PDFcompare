import os
import pandas as pd
from datetime import datetime
from io import BytesIO

def log_to_text(username, product_code, references, match_status, po_number="", so_number=""):
    """Log analysis results to a daily text file"""
    try:
        # Define the directory for text files
        log_dir = r"C:\Users\APP\Desktop\ITL\CSAPP_Logs"
        
        # Create the directory if it doesn't exist
        os.makedirs(log_dir, exist_ok=True)
        
        # Get current date for filename
        current_date = datetime.now().strftime("%Y-%m-%d")
        file_path = os.path.join(log_dir, f"{current_date}.txt")
        
        # Get current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Prepare the log entry
        log_entry = f"{timestamp},{username},{product_code},{references},{po_number},{so_number},{match_status}\n"
        
        # Append to the file
        with open(file_path, "a", encoding="utf-8") as f:
            f.write(log_entry)
        
        return True, f"Report successfully logged to {file_path}"
    except Exception as e:
        return False, f"Error logging to text file: {e}"

def read_log_file_and_convert_to_excel(date_str):
    """Read a log file by date and convert to Excel with separate date and time columns"""
    try:
        # Define the directory for text files
        log_dir = r"C:\Users\APP\Desktop\ITL\CSAPP_Logs"
        file_path = os.path.join(log_dir, f"{date_str}.txt")
        
        # Check if the file exists
        if not os.path.exists(file_path):
            return None, f"Log file for {date_str} not found"
        
        # Read the text file
        with open(file_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        
        # Parse the lines into a DataFrame
        data = []
        for line in lines:
            if line.strip():
                parts = line.strip().split(",")
                if len(parts) >= 7:
                    # Split timestamp into date and time
                    timestamp = parts[0]
                    if " " in timestamp:
                        date_part, time_part = timestamp.split(" ", 1)
                    else:
                        date_part = timestamp
                        time_part = ""
                    
                    data.append({
                        "Date": date_part,
                        "Time": time_part,
                        "User Name": parts[1],
                        "Product code": parts[2],
                        "References": parts[3],
                        "PO Number": parts[4],
                        "SO Number": parts[5],
                        "Status": parts[6]
                    })
        
        # Create DataFrame
        df = pd.DataFrame(data)
        
        # Convert to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Log Data')
        output.seek(0)
        
        return output, f"Successfully converted log file for {date_str}"
    except Exception as e:
        return None, f"Error converting log file: {e}"