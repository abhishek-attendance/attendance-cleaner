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

col1, col2 = st.columns([1, 5])
with col1:
    st.image("logo.png", width=90)
with col2:
    st.markdown("""
        <div style='padding-top:10px'>
        <h1 style='margin-bottom: 0px; color: #2c3e50;'>🧹 Biometric Attendance Cleaner</h1>
        <p style='margin-top: 0px; color: grey;'>Powered by SevenMentor Pvt. Ltd.</p>
        </div>
    """, unsafe_allow_html=True)


st.markdown("This tool lets you upload a messy Excel file and export a clean one with employee info included.")
st.markdown("")

st.markdown("---")
st.markdown(
    "<p style='text-align:center; font-size:0.9em;'>👨‍💻 Developed by <b>Abhishek Wavhal</b></p>",
    unsafe_allow_html=True
)


uploaded_file = st.file_uploader(
    "📂 Upload your raw biometric Excel file (.xlsx / .xlsm)",
    type=["xlsx", "xlsm"]
)

if uploaded_file:
    # -------- Clean the uploaded file --------
    with st.spinner("⚙️ Cleaning in progress…"):
        cleaned_df = clean_workbook(uploaded_file.read())

    if cleaned_df is None:
        st.error("❌ No valid data found in the uploaded file.")
    else:
        st.success(f"✅ Cleaned successfully! Total rows: {len(cleaned_df)}")

        # ---------- Preview ----------
        st.dataframe(cleaned_df.head(50), use_container_width=True)

        # ---------- Prepare download ----------
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            cleaned_df.to_excel(writer, index=False)

        st.download_button(
            label="📥 Download Cleaned File",
            data=out.getvalue(),
            file_name=f"attendance_cleaned_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("⬆️ Upload an Excel file above to begin cleaning.")

# ───────────────────── Footer ─────────────────────
st.markdown("---")
st.markdown(
    "<p style='text-align:center; font-size:0.9em;'>"
    "👨‍💻 Developed by <b>Abhishek Wavhal</b>"
    "</p>",
    unsafe_allow_html=True
)
