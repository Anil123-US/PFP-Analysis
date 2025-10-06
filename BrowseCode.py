import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.title("📊 Project Plan Analysis Tool")

# Upload input files
input_file = st.file_uploader("Upload Latest PFP File", type=["xlsx"])
old_file = st.file_uploader("Upload Old PFP File", type=["xlsx"])

# Session state to persist downloads
if "cleaned_bytes" not in st.session_state:
    st.session_state.cleaned_bytes = None
if "new_bytes" not in st.session_state:
    st.session_state.new_bytes = None
if "today_str" not in st.session_state:
    st.session_state.today_str = None

if input_file and old_file:
    if st.button("Process Files"):
        # === Step 1: Read latest file ===
        df = pd.read_excel(
            input_file,
            dtype={'Project Number': str, 'Employee Name': str},
            engine='openpyxl'
        )

        # Keep only valid rows
        df = df[
            df['Employee Name'].notna()
            & (df['Employee Name'] != '')
            & (df['Employee Name'] != 'Labor Cost, Conversion Employee')
        ]
        df['Employee Name'] = df['Employee Name'].str.strip()
        df['Project Number'] = df['Project Number'].fillna('')

        # Add unique code
        df.insert(0, 'Unique Code', df['Project Number'] + ' - ' + df['Employee Name'])
        df = df.drop_duplicates(subset='Unique Code', keep='first')

        # === Step 6: Compare with old file ===
        old_df = pd.read_excel(
            old_file,
            dtype={'Project Number': str, 'Employee Name': str},
            engine='openpyxl'
        )

        if 'Unique Code' not in old_df.columns:
            old_df['Project Number'] = old_df['Project Number'].fillna('')
            old_df.insert(0, 'Unique Code', old_df['Project Number'] + ' - ' + old_df['Employee Name'])
            old_df = old_df.drop_duplicates(subset='Unique Code', keep='first')

        # Find new records
        new_codes = set(df['Unique Code']) - set(old_df['Unique Code'])
        new_rows = df[df['Unique Code'].isin(new_codes)].copy()

        # === Format dates (DD-MM-YYYY) ===
        for col in df.columns:
            if "Date" in col or "date" in col:
                try:
                    df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime("%d-%m-%Y")
                    new_rows[col] = pd.to_datetime(new_rows[col], errors='coerce').dt.strftime("%d-%m-%Y")
                except Exception:
                    pass  # ignore columns that aren't real dates

        # Today string for filenames (DDMMYYYY)
        today_str = datetime.today().strftime("%d%m%Y")

        # === Save cleaned file (to memory) ===
        cleaned_buffer = io.BytesIO()
        df.to_excel(cleaned_buffer, index=False, engine='openpyxl')
        st.session_state.cleaned_bytes = cleaned_buffer.getvalue()

        # === Save new records file (to memory) ===
        new_buffer = io.BytesIO()
        new_rows.to_excel(new_buffer, index=False, engine='openpyxl')
        st.session_state.new_bytes = new_buffer.getvalue()

        st.session_state.today_str = today_str
        st.success("✅ Processing complete! Files are ready for download.")

# === Always show download buttons if files exist ===
if st.session_state.cleaned_bytes:
    st.download_button(
        "⬇️ Download Cleaned PFP",
        st.session_state.cleaned_bytes,
        file_name=f"Project Plan Analysis-continuous-{st.session_state.today_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if st.session_state.new_bytes:
    st.download_button(
        "⬇️ Download NEW PFP",
        st.session_state.new_bytes,
        file_name=f"NEW PFP-{st.session_state.today_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
