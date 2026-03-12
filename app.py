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
    """Processes an uploaded Excel file in memory."""
    # Load the uploaded file using BytesIO
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
st.markdown("""
This tool checks character limits for localization files. 
**Privacy Note:** Files are processed in your browser's session memory and are deleted immediately after use.
""")

# Sidebar for settings
with st.sidebar:
    st.header("Settings")
    
    # Language Dropdown
    languages = ["de-DE", "es-ES", "es-MX", "fr-FR", "hi-IN", "it-IT", "ja-JP", "pt-BR"]
    lang_code = st.selectbox("Select Language Code", options=languages, index=0)
    
    st.info(f"The report will be generated for: **{lang_code}**")

# File Uploader
uploaded_files = st.file_uploader(
    "Upload Excel Files (.xlsx, .xlsm)", 
    type=["xlsx", "xlsm"], 
    accept_multiple_files=True
)

if uploaded_files:
    all_violations = []
    
    # Progress Bar
    progress_bar = st.progress(0)
    for idx, uploaded_file in enumerate(uploaded_files):
        violations = process_excel(uploaded_file, lang_code)
        all_violations.extend(violations)
        progress_bar.progress((idx + 1) / len(uploaded_files))

    if all_violations:
        df = pd.DataFrame(all_violations)
        
        # UI Feedback
        st.error(f"⚠️ Found {len(df)} violations across {len(uploaded_files)} files.")
        
        # Display Preview
        st.subheader("Violation Preview")
        st.dataframe(df, use_container_width=True)

        # Generate Report in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Violations")
            
            # Simple formatting for the output file
            ws = writer.sheets["Violations"]
            for cell in ws[1]: # Bold headers
                cell.font = cell.font.copy(bold=True)

        # Download Button
        st.download_button(
            label="📥 Download Full Report (.xlsx)",
            data=output.getvalue(),
            file_name=f"length_violations_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.balloons()
        st.success("✅ No violations found! All files are within character limits.")
