import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Alignment

# --- CONFIG ---
THRESHOLDS = {"F": 100, "G": 41, "H": 41, "I": 41, "J": 41, "K": 41}

st.set_page_config(page_title="Amazon Music Length Validator", layout="wide")

def excel_len(v) -> int:
    return len(str(v)) if v is not None else 0

def process_excel(file, lang):
    # Load the uploaded file into memory
    wb = load_workbook(file, data_only=True)
    ws = wb.worksheets[0]
    
    violations = []
    target_letters = list(THRESHOLDS.keys())
    col_indexes = [column_index_from_string(l) for l in target_letters]

    for r in range(1, ws.max_row + 1):
        for c, col_letter in zip(col_indexes, target_letters):
            cell = ws.cell(row=r, column=c)
            actual = excel_len(cell.value)
            needed = THRESHOLDS[col_letter]
            if actual > needed:
                violations.append({
                    "language": lang,
                    "file": file.name,
                    "row": r,
                    "column": col_letter,
                    "needed_length": needed,
                    "actual_length": actual,
                    "over_by": actual - needed,
                    "value": cell.value
                })
    return violations

# --- UI ---
st.title("🎵 Amazon Music Length Validator")
st.info("Files are processed in memory and are not stored permanently.")

# Language Input
lang_code = st.text_input("Enter Language Code (e.g., de-DE, fr-FR)", "en-US")

# File Uploader
uploaded_files = st.file_uploader("Upload Excel Files", type=["xlsx", "xlsm"], accept_multiple_files=True)

if uploaded_files:
    all_violations = []
    for uploaded_file in uploaded_files:
        with st.spinner(f"Processing {uploaded_file.name}..."):
            violations = process_excel(uploaded_file, lang_code)
            all_violations.extend(violations)

    if all_violations:
        df = pd.DataFrame(all_violations)
        st.success(f"Found {len(df)} violations!")
        
        # Display Preview
        st.dataframe(df)

        # Create Downloadable Excel in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Violations")
        
        st.download_button(
            label="📥 Download Report",
            data=output.getvalue(),
            file_name=f"violations_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.balloons()
        st.success("No violations found!")