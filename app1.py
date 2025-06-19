import streamlit as st
import os
import glob
import pandas as pd

def get_valid_excel_files(folder_path):
    excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    valid_excel_files = []
    for file in excel_files:
        try:
            pd.ExcelFile(file)
            valid_excel_files.append(file)
        except Exception as e:
            st.warning(f"⚠️ Skipped invalid Excel file: {file} ({e})")
    return valid_excel_files

def clean_feedback_excel(input_path):
    xls = pd.ExcelFile(input_path)
    all_cleaned_sheets = {}

    for sheet_name in xls.sheet_names:
        st.text(f"🔍 Processing sheet '{sheet_name}' in '{os.path.basename(input_path)}'")
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
            st.warning(f"⚠️ No data blocks found in sheet '{sheet_name}'")
            all_cleaned_sheets[sheet_name] = pd.DataFrame()

    return all_cleaned_sheets

def process_and_merge_all_files(folder_path, merged_output_path):
    input_files = get_valid_excel_files(folder_path)
    if not input_files:
        st.error("❌ No valid Excel files found in the folder.")
        return False

    with pd.ExcelWriter(merged_output_path) as writer:
        for file in input_files:
            cleaned_sheets = clean_feedback_excel(file)
            file_base = os.path.splitext(os.path.basename(file))[0][:20]

            for sheet_name, df in cleaned_sheets.items():
                combined_sheet_name = f"{file_base}_{sheet_name}"[:31]
                df.to_excel(writer, sheet_name=combined_sheet_name, index=False)

    st.success(f"✅ Cleaned and combined {len(input_files)} files into:\n{merged_output_path}")
    return True

# ---- Streamlit app UI ----

st.title("Excel Feedback Cleaner & Merger")

folder_path = st.text_input("Enter folder path containing Excel files:", "")

if folder_path:
    processed_folder = os.path.join(folder_path, "Processed")
    os.makedirs(processed_folder, exist_ok=True)
    merged_output_path = os.path.join(processed_folder, "all_cleaned_feedback.xlsx")

    if st.button("Run Cleaning and Merge"):
        with st.spinner("Processing..."):
            success = process_and_merge_all_files(folder_path, merged_output_path)
        if success:
            with open(merged_output_path, "rb") as f:
                st.download_button(
                    label="Download Cleaned & Combined Excel",
                    data=f,
                    file_name="all_cleaned_feedback.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
