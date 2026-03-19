import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime

# App Configuration
st.set_page_config(page_title="Ward Master Splitter", page_icon="🏢", layout="wide")

st.title("Customized Ward-wise Excel Splitter 📊")
st.write("Upload your Excel file, select Year/Month, and download Ward-wise consolidated files.")

# --- UI for Year and Month Selection ---
col1, col2 = st.columns(2)
with col1:
    years = [str(year) for year in range(2024, 2031)]
    selected_year = st.selectbox("Select Year:", years, index=2) # Default 2026

with col2:
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    selected_month = st.selectbox("Select Month:", months, index=1) # Default Feb

# File uploader
uploaded_file = st.file_uploader("Upload your Master Excel file (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        st.success(f"File loaded with {len(sheet_names)} sheets.")
        
        all_wards = set()
        data_dict = {} 
        column_headers_dict = {} # To store headers if data is empty

        with st.spinner('Analyzing Wards and Sheets...'):
            for sheet in sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet)
                
                # Store original headers (Col E to S) for empty sheets
                if len(df.columns) >= 19:
                    headers = df.iloc[:, 4:19].columns.tolist()
                    column_headers_dict[sheet] = headers

                if 'Ward' in df.columns:
                    df['Ward'] = df['Ward'].astype(str).str.strip()
                    unique_in_sheet = df['Ward'].unique().tolist()
                    all_wards.update(unique_in_sheet)
                    data_dict[sheet] = df
                else:
                    st.warning(f"Column 'Ward' not found in sheet: {sheet}")

        # Clean ward list
        all_wards = [w for w in all_wards if w not in ['nan', 'None', '', 'nan']]
        
        if not all_wards:
            st.error("No Wards found! Please check the 'Ward' column in your Excel.")
        else:
            st.info(f"Detected {len(all_wards)} Wards. Click below to generate files.")
            
            if st.button("🚀 Generate ZIP with Custom Naming"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for ward in sorted(all_wards):
                        ward_buffer = io.BytesIO()
                        
                        with pd.ExcelWriter(ward_buffer, engine='openpyxl') as writer:
                            for sheet_name in sheet_names:
                                if sheet_name in data_dict:
                                    df = data_dict[sheet_name]
                                    ward_df = df[df['Ward'] == ward].copy()
                                    
                                    # Select Cols E to S
                                    if len(ward_df.columns) >= 19:
                                        ward_df = ward_df.iloc[:, 4:19]
                                    
                                    # If no data for this ward, keep only headers
                                    if ward_df.empty:
                                        ward_df = pd.DataFrame(columns=column_headers_dict.get(sheet_name, []))
                                    
                                    # Add Sr. No. (even if empty, it's better to have the column)
                                    ward_df.insert(0, 'Sr. No.', range(1, len(ward_df) + 1))
                                    
                                    ward_df.to_excel(writer, index=False, sheet_name=sheet_name)
                        
                        # CUSTOM FILENAME LOGIC
                        # Example: A_Ward_2026_Feb_Month Lab Confirmed Line List of Monsoon Related Diseases
                        custom_filename = f"{ward}_Ward_{selected_year}_{selected_month}_Month Lab Confirmed Line List of Monsoon Related Diseases.xlsx"
                        zip_file.writestr(custom_filename, ward_buffer.getvalue())
                
                st.success("All Ward files are ready!")
                
                st.download_button(
                    label="📥 Download ZIP Folder",
                    data=zip_buffer.getvalue(),
                    file_name=f"Wards_Data_{selected_month}_{selected_year}.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Something went wrong: {e}")
