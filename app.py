import streamlit as st
import pandas as pd
import math

# ==============================
# KONFIGURASI HALAMAN
# ==============================
st.set_page_config(page_title="ETMAL Calculator", layout="wide")
st.title("🚢 ETMAL & Invoice Calculator")

st.write(
    "Upload file Excel untuk menghitung "
    "**ETMAL Charge** dan **Invoice (USD)** secara otomatis."
)

uploaded_file = st.file_uploader(
    "📤 Upload file Excel (.xlsx)",
    type=["xlsx"]
)

# ==============================
# FUNGSI HITUNG ETMAL
# ==============================
def hitung_etmal(jam):
    if pd.isna(jam) or jam <= 0:
        return 0
    return min(math.ceil(jam / 6) * 0.25, 2)

# ==============================
# PROSES FILE
# ==============================
if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_aktif = xls.sheet_names[0]
        st.info(f"📄 Sheet digunakan: **{sheet_aktif}**")

        df = pd.read_excel(xls, sheet_name=sheet_aktif)

        # ==============================
        # KONVERSI KOLOM JAM (KRITIS)
        # ==============================
        jam_berthing = pd.to_numeric(
            df.iloc[:, 119],  # DP
