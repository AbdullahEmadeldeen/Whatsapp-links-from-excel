import streamlit as st
import pandas as pd
from re import search as re_search
from io import BytesIO

# -----------------------------
# Helper function to extract phone numbers
# -----------------------------
def extract_phone_number(text):
    match = re_search(r'01\d{9}', str(text))
    if match:
        return '+20' + match.group()
    return None

# -----------------------------
# Function to process Excel sheet
# -----------------------------
def process_excel(file, sheet_name):
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, dtype=str)

        if df.shape[1] < 2:
            st.error("❌ Selected sheet does not have a second column!")
            return None

        # Extract phone numbers from column 2
        phone_numbers = df.iloc[:, 1].dropna().astype(str).apply(extract_phone_number)
        phone_numbers = phone_numbers.dropna().unique()

        # Create new DataFrame
        new_data = pd.DataFrame({'Phone Number': phone_numbers})
        new_data['WhatsApp Link'] = new_data['Phone Number'].apply(lambda x: f'https://wa.me/{x}' if x else None)
        
        return new_data

    except Exception as e:
        st.error(f"⚠️ Error: {e}")
        return None

# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Excel Phone Extractor", page_icon="📱", layout="centered")

st.title("📱 Excel Phone Extractor")
st.write("Upload an Excel file, select a sheet, and extract phone numbers with WhatsApp links.")

# Upload Excel file
uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Get all sheet names
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("🧾 Select a sheet", xls.sheet_names)

        if st.button("Process Sheet"):
            result_df = process_excel(uploaded_file, sheet_name)

            if result_df is not None and not result_df.empty:
                st.success("✅ File processed successfully!")
                st.dataframe(result_df)

                # Convert to downloadable Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    result_df.to_excel(writer, index=False)
                st.download_button(
                    label="💾 Download Processed Excel",
                    data=output.getvalue(),
                    file_name=f"{sheet_name}_processed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No valid phone numbers found in the selected sheet.")

    except Exception as e:
        st.error(f"⚠️ Could not read Excel file: {e}")
