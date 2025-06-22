import streamlit as st
import pandas as pd
from pathlib import Path

# Page Configuration
st.set_page_config(
    page_title="NMITE Excel Feedback Cleaner Merger",
    page_icon="Nmite_Logo.jpg",  # Logo appears in the browser tab
    layout="wide"
)

# Custom CSS Styling (colors and footer)
custom_css = """
<style>
    .stApp {
        background-color: #e6f2ff;
        color: #003366;
        font-family: "Segoe UI", sans-serif;
    }
    h1, h2, h3, h4, h5, h6, p, label, .css-1d391kg {
        color: #003366 !important;
    }
    .footer-container {
        position: fixed;
        bottom: 10px;
        left: 0;
        right: 0;
        text-align: center;
        font-size: 16px;
        font-weight: bold;
        color: #003366;
    }
</style>
<div class="footer-container">
    Developed by Dr James Atuonwu for NMITE | james.atuonwu@nmite.ac.uk
</div>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# Title Section with Logo and Header Text
col1, col2 = st.columns([1, 6])
with col1:
    st.image("Nmite_Logo.jpg", width=100)
with col2:
    st.title("📊 NMITE Excel Feedback Cleaner Merger")
st.markdown("Upload multiple Excel feedback files, clean and merge them effortlessly.")

# Core Cleaning Function
def clean_feedback_excel_from_file(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    all_cleaned_sheets = {}

    for sheet_name in xls.sheet_names:
        st.text(f"🔍 Processing: {uploaded_file.name}")
        df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
        df = df.iloc[:, 9:].copy()
        n_cols = df.shape[1]
        i = 0
        all_data_blocks = []

        while i < n_cols:
            question_header_col = df.columns[i]
            col_data = df.iloc[:, i]

            if col_data.dropna().empty:
                question = question_header_col
                data_cols = {}
                j = i + 1

                while j < n_cols:
                    module_col_header = df.columns[j]
                    module_col_data = df.iloc[:, j]

                    if module_col_data.dropna().empty:
                        break

                    data_cols[module_col_header] = module_col_data.reset_index(drop=True)

                    for extra_col_offset in [1, 2]:
                        extra_col_idx = j + extra_col_offset
                        if extra_col_idx < n_cols:
                            extra_col_header = df.columns[extra_col_idx]
                            extra_col_data = df.iloc[:, extra_col_idx]

                            if not extra_col_data.dropna().empty:
                                data_cols[extra_col_header] = extra_col_data.reset_index(drop=True)
                    j += 3

                if data_cols:
                    question_df = pd.DataFrame(data_cols)
                    question_df.columns = pd.MultiIndex.from_product([[question], question_df.columns])
                    all_data_blocks.append(question_df.reset_index(drop=True))

                i = j
            else:
                i += 1

        if all_data_blocks:
            final_df = pd.concat(all_data_blocks, axis=1)
            final_df.columns = [' - '.join(col).strip() for col in final_df.columns]
            all_cleaned_sheets[sheet_name] = final_df
        else:
            st.warning(f"⚠️ No data blocks found in {uploaded_file.name}")
            all_cleaned_sheets[sheet_name] = pd.DataFrame()

    return all_cleaned_sheets

# Merge and Save Function
def process_and_merge_uploaded_files(uploaded_files):
    if not uploaded_files:
        st.error("❌ No files uploaded.")
        return None

    output_file = "all_cleaned_feedback.xlsx"
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for uploaded_file in uploaded_files:
            cleaned_sheets = clean_feedback_excel_from_file(uploaded_file)
            file_base = uploaded_file.name.rsplit('.', 1)[0][:31]
            for _, df in cleaned_sheets.items():
                df.to_excel(writer, sheet_name=file_base, index=False)

    return output_file

# ---- Streamlit UI ----
uploaded_files = st.file_uploader("📤 Upload Excel feedback files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Run Cleaning and Merge"):
        with st.spinner("⏳ Processing..."):
            output_path = process_and_merge_uploaded_files(uploaded_files)

        if output_path:
            st.success("✅ Processing Complete!")
            with open(output_path, "rb") as f:
                st.download_button(
                    label="⬇️ Download Cleaned & Combined Excel",
                    data=f,
                    file_name="all_cleaned_feedback.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
