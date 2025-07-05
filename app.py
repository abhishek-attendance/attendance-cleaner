import streamlit as st
import pandas as pd
import io
from datetime import date

required_columns = [
    "Date", "ArrTim", "LateHrs", "DepTim", "EarlHrs", "WrkHrs", "OvTim",
    "Present", "Absent", "Paid_Lv", "UnPaidLv", "PrsAbs", "Remarks"
]

def find_header_row(df):
    for i, row in df.iterrows():
        if any(str(c).strip().lower() in {c.lower() for c in required_columns} for c in row):
            return i
    return None

def extract_emp_metadata(df_raw, lookahead=10):
    empcode = empname = None
    for _, row in df_raw.head(lookahead).iterrows():
        vals = [str(v).strip() for v in row if pd.notnull(v)]
        for i, v in enumerate(vals):
            vl = v.lower()
            if vl == "empcode" and i + 1 < len(vals):
                empcode = vals[i + 1]
            if vl == "name" and i + 1 < len(vals):
                empname = vals[i + 1]
    return empcode, empname

@st.cache_data(show_spinner=False)
def clean_workbook(raw_bytes):
    xls = pd.ExcelFile(io.BytesIO(raw_bytes))
    frames = []
    for sheet in xls.sheet_names:
        df_raw = xls.parse(sheet, header=None)
        empcode, empname = extract_emp_metadata(df_raw)
        hdr = find_header_row(df_raw)
        if hdr is None:
            continue
        df = xls.parse(sheet, header=hdr)
        df = df[[c for c in required_columns if c in df.columns]]
        df = df[~df.iloc[:, 0].astype(str).str.strip().str.lower().eq("total for :")]
        df["Empcode"] = empcode
        df["EmpName"] = empname
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else None

# ───────────────────── UI ─────────────────────
st.set_page_config(page_title="Attendance Cleaner", page_icon="🧹", layout="centered")

col1, col2 = st.columns([1, 3])
with col1:
    st.image("logo.png", width=110)
with col2:
    st.markdown(
        '<div style="font-size: 2em; font-weight: bold; color: #4CAF50;">'
        '🧹 Biometric Attendance Cleaner</div>',
        unsafe_allow_html=True
    )

st.markdown("This tool lets you upload a messy Excel file and export a clean one with employee info included.")
st.markdown("")

uploaded_file = st.file_uploader("📂 Upload your raw biometric Excel file (.xlsx)", type=["xlsx", "xlsm"])

if uploaded_file:
    with st.spinner("⚙️ Cleaning in progress..."):
        cleaned_df = clean_workbook(uploaded_file.read())

    if cleaned_df is None:
        st.error("❌ No valid data found in the uploaded file.")
    else:
