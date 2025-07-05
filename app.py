import streamlit as st
import pandas as pd
import io
from datetime import datetime

# Helper functions
required_columns = [
    "Date", "ArrTim", "LateHrs", "DepTim", "EarlHrs", "WrkHrs", "OvTim",
    "Present", "Absent", "Paid_Lv", "UnPaidLv", "PrsAbs", "Remarks"
]

def find_header_row(df):
    for i, row in df.iterrows():
        if any(str(cell).strip().lower() in [c.lower() for c in required_columns]
               for cell in row):
            return i
    return None

def extract_emp_metadata(df_raw, lookahead=10):
    empcode = empname = None
    for _, row in df_raw.head(lookahead).iterrows():
        vals = [str(v).strip() for v in row if pd.notnull(v)]
        for idx, v in enumerate(vals):
            v_low = v.lower()
            if v_low == "empcode" and idx + 1 < len(vals):
                empcode = vals[idx + 1]
            if v_low == "name" and idx + 1 < len(vals):
                empname = vals[idx + 1]
    return empcode, empname

def clean_workbook(uploaded_bytes):
    xls = pd.ExcelFile(uploaded_bytes)
    cleaned = []
    for sheet in xls.sheet_names:
        df_raw = xls.parse(sheet, header=None)
        empcode, empname = extract_emp_metadata(df_raw)
        hdr = find_header_row(df_raw)
        if hdr is None:
            continue
        df = xls.parse(sheet, header=hdr)
        df = df[[c for c in required_columns if c in df.columns]]
        df = df[~df.iloc[:,0].astype(str).str.strip().str.lower().eq("total for :")]
        df["Empcode"] = empcode
        df["EmpName"] = empname
        cleaned.append(df)
    return pd.concat(cleaned, ignore_index=True) if cleaned else None

# Streamlit app
st.title("📊 Attendance Cleaner")

uploaded = st.file_uploader("Upload raw biometric Excel file", type=["xlsx", "xlsm"])

if uploaded:
    with st.spinner("Cleaning… please wait..."):
        cleaned_df = clean_workbook(uploaded)
    if cleaned_df is None:
        st.error("❌ No valid sheets found.")
    else:
        st.success(f"✅ Cleaned {len(cleaned_df)} rows!")
        st.dataframe(cleaned_df.head())

        # Download link
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            cleaned_df.to_excel(writer, index=False)
        st.download_button(
            label="📥 Download cleaned file",
            data=buf.getvalue(),
            file_name=f"attendance_cleaned_{datetime.today().date()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
