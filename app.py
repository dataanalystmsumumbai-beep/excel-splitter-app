import streamlit as st
import pandas as pd
import io
import zipfile

# App Configuration
st.set_page_config(page_title="Ward Master Splitter", page_icon="🏢", layout="wide")

st.title("Ward-wise Excel Splitter 📊")
st.write("Upload your Excel, select Year/Month, and download formatted Ward-wise files.")

# --- UI for Year and Month Selection ---
col1, col2 = st.columns(2)

with col1:
    years = [str(year) for year in range(2024, 2031)]
    selected_year = st.selectbox("Select Year:", years, index=2)

with col2:
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    selected_month = st.selectbox("Select Month:", months, index=1)

# File uploader
uploaded_file = st.file_uploader("Upload your Master Excel file (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        st.success(f"File loaded with {len(sheet_names)} sheets.")

        all_wards = set()
        data_dict = {}
        column_headers_dict = {}

        with st.spinner('Analyzing Wards and Formatting Dates...'):
            for sheet in sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet)

                # ✅ Clean Ward column properly
                if 'Ward' in df.columns:
                    df['Ward'] = df['Ward'].fillna('').astype(str).str.strip()

                # ✅ Date Formatting
                for col in ['Reporting Date', 'Date Of Onset']:
                    if col in df.columns:
                        df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d-%b-%y')

                # ✅ Store headers (Col E to S)
                if len(df.columns) >= 19:
                    column_headers_dict[sheet] = df.iloc[:, 4:19].columns.tolist()

                # ✅ Collect wards
                if 'Ward' in df.columns:
                    valid_wards = df['Ward'].unique().tolist()
                    all_wards.update(valid_wards)
                    data_dict[sheet] = df

        # ✅ Clean ward list properly
        all_wards = [
            w for w in all_wards
            if w and str(w).strip().lower() not in ['nan', 'none', 'null']
        ]

        if not all_wards:
            st.error("No Wards found! Please check the 'Ward' column.")
        else:
            st.info(f"Detected {len(all_wards)} Wards.")

            if st.button("🚀 Generate Formatted ZIP"):

                zip_buffer = io.BytesIO()

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

                    # ✅ Safe sorting
                    sorted_wards = sorted(
                        all_wards,
                        key=lambda x: int(x) if str(x).isdigit() else str(x)
                    )

                    for ward in sorted_wards:
                        ward_buffer = io.BytesIO()

                        with pd.ExcelWriter(ward_buffer, engine='openpyxl') as writer:

                            for sheet_name in sheet_names:
                                if sheet_name in data_dict:
                                    df = data_dict[sheet_name]

                                    ward_df = df[df['Ward'] == ward].copy()

                                    # ✅ Select columns E to S
                                    if len(ward_df.columns) >= 19:
                                        ward_df = ward_df.iloc[:, 4:19]

                                    # ✅ Empty sheet handling
                                    if ward_df.empty:
                                        ward_df = pd.DataFrame(
                                            columns=column_headers_dict.get(sheet_name, [])
                                        )

                                    # ✅ Add Serial Number
                                    ward_df.insert(0, 'Sr. No.', range(1, len(ward_df) + 1))

                                    ward_df.to_excel(writer, index=False, sheet_name=sheet_name)

                        # ✅ File naming
                        filename = f"{ward}_Ward_{selected_month}_{selected_year}_Lab Confirmed Line List.xlsx"

                        zip_file.writestr(filename, ward_buffer.getvalue())

                st.success("All files formatted and ready!")

                st.download_button(
                    label="📥 Download ZIP Folder",
                    data=zip_buffer.getvalue(),
                    file_name=f"Wards_Data_{selected_month}_{selected_year}.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Error: {e}")
