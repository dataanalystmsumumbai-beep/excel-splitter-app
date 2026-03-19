import streamlit as st
import pandas as pd
import io
import zipfile

# App Configuration
st.set_page_config(page_title="Ward-wise Excel Splitter", page_icon="🏢", layout="wide")

st.title("Excel Ward-wise Master Splitter 📊")
st.write("Upload your Excel file to split data by Ward into a single ZIP file.")

# File uploader
uploaded_file = st.file_uploader("Upload your Master Excel file (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # Load the Excel file to get sheet names
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        st.success(f"File loaded with {len(sheet_names)} sheets.")
        
        # 1. Identify all unique Wards across all sheets
        all_wards = set()
        data_dict = {} # To store data frames for each sheet
        
        with st.spinner('Processing sheets and identifying Wards...'):
            for sheet in sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet)
                
                # Check if 'Ward' column exists
                if 'Ward' in df.columns:
                    # Clean ward names (remove spaces)
                    df['Ward'] = df['Ward'].astype(str).str.strip()
                    unique_in_sheet = df['Ward'].unique().tolist()
                    all_wards.update(unique_in_sheet)
                    data_dict[sheet] = df
                else:
                    st.warning(f"Column 'Ward' not found in sheet: {sheet}. Skipping this sheet.")

        # Remove any 'nan' or empty ward names from the list
        all_wards = [w for w in all_wards if w != 'nan' and w != 'None' and w != '']
        
        if not all_wards:
            st.error("No Wards found in the 'Ward' column across any sheets.")
        else:
            st.write(f"Found **{len(all_wards)}** unique Wards.")
            
            # Button to process and Create ZIP
            if st.button("🚀 Generate & Download All Ward Files (ZIP)"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for ward in sorted(all_wards):
                        # Create an Excel file for each Ward in memory
                        ward_buffer = io.BytesIO()
                        
                        with pd.ExcelWriter(ward_buffer, engine='openpyxl') as writer:
                            for sheet_name, df in data_dict.items():
                                # Filter data for this specific ward
                                ward_df = df[df['Ward'] == ward].copy()
                                
                                if not ward_df.empty:
                                    # 2. Select Columns E to S (Reporting Date to Facility Name Lform)
                                    # Note: Python index starts at 0, so Col E is 4 and Col S is 18
                                    # We use iloc to get columns from index 4 to 19 (19 is exclusive)
                                    if len(ward_df.columns) >= 19:
                                        ward_df = ward_df.iloc[:, 4:19]
                                    
                                    # 3. Add Sr. No. at the beginning
                                    ward_df.insert(0, 'Sr. No.', range(1, len(ward_df) + 1))
                                    
                                    # Write to sheet
                                    ward_df.to_excel(writer, index=False, sheet_name=sheet_name)
                        
                        # Add this ward's Excel file to the ZIP
                        zip_file.writestr(f"{ward}_Consolidated_Data.xlsx", ward_buffer.getvalue())
                
                st.success("All files processed successfully!")
                
                # Download button for the ZIP file
                st.download_button(
                    label="📥 Download All Wards (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="All_Wards_Data.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")
