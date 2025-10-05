# streamlit_app.py
import re
from io import BytesIO
import pandas as pd
import streamlit as st

# -----------------------------
# Page config
# -----------------------------
st.set_page_config(
    page_title="Excel Phone Extractor → WhatsApp Links",
    page_icon="📇",
    layout="wide",
)

st.title("📇 Excel Phone Extractor → WhatsApp Links")
st.caption("Upload an Excel or enter numbers manually. Extract Egyptian mobile numbers → build WhatsApp links → download a clean Excel.")

# -----------------------------
# Helpers: numbers & links
# -----------------------------
EG_MOBILE_REGEX = re.compile(r'(?:\+?20)?0?1\d{9}')  # handles +20 / 20 / 01 prefixes

def find_egypt_mobile(text: str) -> str | None:
    """
    Find an Egyptian mobile number and normalize to digits-only international format for wa.me:
      returns: 201XXXXXXXXX (digits only) or None if invalid.
    Display format will be +201XXXXXXXXX; link uses https://wa.me/201XXXXXXXXX
    """
    if text is None:
        return None
    s = str(text)

    m = EG_MOBILE_REGEX.search(s)
    if not m:
        return None

    digits = re.sub(r"\D", "", m.group())

    # Normalize to international digits-only:
    if digits.startswith("0") and len(digits) == 11:
        digits = "20" + digits[1:]
    elif digits.startswith("20") and len(digits) == 12:
        pass
    elif digits.startswith("1") and len(digits) == 10:
        digits = "20" + digits
    else:
        return None

    if not (digits.startswith("201") and len(digits) == 12):
        return None

    return digits  # e.g., 2010XXXXXXXX


def build_output_df(raw_series: pd.Series) -> pd.DataFrame:
    """
    From a pandas Series (raw values), extract valid Egyptian numbers and build:
    - Phone Number (display): +201XXXXXXXXX
    - WhatsApp Link: https://wa.me/201XXXXXXXXX
    - Completed: False
    Deduplicates by Phone Number.
    """
    rows = []
    for val in raw_series:
        digits = find_egypt_mobile(val)
        if digits:
            rows.append({
                "Phone Number": f"+{digits}",
                "WhatsApp Link": f"https://wa.me/{digits}",
                "Completed": False
            })

    df = pd.DataFrame(rows, columns=["Phone Number", "WhatsApp Link", "Completed"])
    if not df.empty:
        df = df.drop_duplicates(subset=["Phone Number"]).reset_index(drop=True)
    return df


def dataframe_to_excel_bytes(df: pd.DataFrame, make_clickable=False) -> bytes:
    """
    Build an in-memory XLSX. Optionally replace WhatsApp Link with Excel HYPERLINK formula.
    Only two columns are exported: Phone Number, WhatsApp Link.
    """
    out_df = df.copy()
    out_df = out_df[["Phone Number", "WhatsApp Link"]]

    if make_clickable:
        out_df["WhatsApp Link"] = out_df["WhatsApp Link"].apply(
            lambda url: f'=HYPERLINK("{url}", "Open WhatsApp")'
        )

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="WhatsApp")
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------
# UI - Tabs for flows
# -----------------------------
tab_upload, tab_manual = st.tabs(["📤 Upload Excel", "📝 Manual Entry"])

# ==========================================
# Tab 1: Upload Excel
# ==========================================
with tab_upload:
    st.subheader("Upload Excel")
    file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

    if file:
        try:
            xls = pd.ExcelFile(file)
            sheet = st.selectbox("Select sheet", xls.sheet_names, index=0)
            df = pd.read_excel(xls, sheet_name=sheet, dtype=str)

            st.write("Preview:")
            st.dataframe(df.head(20), width="stretch")

            if df.shape[1] < 2:
                st.warning("The original script expected **phone numbers in the 2nd column**. Select the correct column below.")
            default_col_idx = 1 if df.shape[1] >= 2 else 0
            col_name = st.selectbox("Which column contains phone numbers?", df.columns, index=default_col_idx)

            # Process
            result_df = build_output_df(df[col_name])

            c1 = st.columns(1)[0]
            with c1:
                st.metric("Valid numbers found", len(result_df))

            st.divider()
            st.subheader("Result (editable)")
            st.caption("Tick ✅ Completed to track progress. Links are clickable.")

            edited_df = st.data_editor(
                result_df,
                width="stretch",
                num_rows="fixed",
                hide_index=True,
                column_config={
                    "WhatsApp Link": st.column_config.LinkColumn("WhatsApp Link", help="Open WhatsApp chat"),
                    "Completed": st.column_config.CheckboxColumn("Completed", help="Mark as processed"),
                },
            )

            st.divider()
            st.subheader("Download")

            make_clickable = st.toggle("Make Excel links clickable (Excel HYPERLINK formula)", value=True)
            excel_bytes = dataframe_to_excel_bytes(edited_df, make_clickable=make_clickable)

            st.download_button(
                label="⬇️ Download processed Excel (2 columns)",
                data=excel_bytes,
                file_name="processed_whatsapp.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Error reading file: {e}")

# ==========================================
# Tab 2: Manual Entry
# ==========================================
with tab_manual:
    st.subheader("Manual Entry")
    st.caption("Paste or type Egyptian mobile numbers (11 digits starting with 01). Mixed text is okay; the number will be extracted.")

    placeholder = "e.g.\n01012345678\n01198765432\nText with 015XXXXXXX mixed in"
    raw_text = st.text_area("Numbers (one per line or mixed in text)", height=180, placeholder=placeholder)

    col_a, col_b = st.columns([1, 2])
    with col_a:
        parse_btn = st.button("Parse numbers")
    with col_b:
        st.caption("Click **Parse numbers** to extract, clean, and build WhatsApp links.")

    if parse_btn and raw_text.strip():
        lines = [ln for ln in raw_text.splitlines() if ln.strip()]
        series = pd.Series(lines, dtype="string")
        manual_df = build_output_df(series)

        if manual_df.empty:
            st.warning("No valid Egyptian mobile numbers found.")
        else:
            st.success(f"Found {len(manual_df)} valid numbers.")

            st.divider()
            st.subheader("Result (add/edit & mark completed)")
            st.caption("Add more rows if you like. Use the link to open WhatsApp.")

            editable_df = st.data_editor(
                manual_df,
                width="stretch",
                num_rows="dynamic",
                hide_index=True,
                column_config={
                    "WhatsApp Link": st.column_config.LinkColumn("WhatsApp Link", help="Open WhatsApp chat"),
                    "Completed": st.column_config.CheckboxColumn("Completed", help="Mark as processed"),
                },
            )

            st.divider()
            st.subheader("Download")
            make_clickable_2 = st.toggle("Make Excel links clickable (Excel HYPERLINK formula)", key="clickable2", value=True)
            excel_bytes_2 = dataframe_to_excel_bytes(editable_df, make_clickable=make_clickable_2)

            st.download_button(
                label="⬇️ Download processed Excel (2 columns)",
                data=excel_bytes_2,
                file_name="processed_whatsapp_manual.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with st.expander("ℹ️ What numbers are supported?"):
        st.markdown(
            """
            - **Egyptian mobile numbers** that look like `01xxxxxxxxx` (010/011/012/015 + 8 digits).
            - Accepted inputs: `01...`, `+201...`, or `201...` embedded in text.
            - Output:
              - **Display:** `+201xxxxxxxxx`
              - **WhatsApp:** `https://wa.me/201xxxxxxxxx` (digits only — no `+`).
            """
        )

