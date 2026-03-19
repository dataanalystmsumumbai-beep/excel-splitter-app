import streamlit as st
import pandas as pd
import io

# App Configuration
st.set_page_config(page_title="Excel Splitter", page_icon="📊")

st.title("Excel File Splitter App 📊")
st.write("Upload your Excel file containing multiple sheets and extract the specific sheet you need.")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # Read all sheets from the Excel file
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        st.success("File uploaded successfully!")
        
        # Selectbox to choose the sheet
        selected_sheet = st.selectbox("Which sheet would you like to extract?", sheet_names)
        
        if selected_sheet:
            # Read and display the selected sheet
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            st.write(f"Preview of **{selected_sheet}** (First 5 rows):")
            st.dataframe(df.head()) 
            
            # Create a buffer to save the new Excel file
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name=selected_sheet)
            
            # Download button
            st.download_button(
                label=f"📥 Download '{selected_sheet}' Data",
                data=buffer.getvalue(),
                file_name=f"{selected_sheet}_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"An error occurred while reading the file: {e}")
