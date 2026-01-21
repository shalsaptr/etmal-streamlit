import streamlit as st
import pandas as pd
import math

# ==============================
# KONFIGURASI HALAMAN
# ==============================
st.set_page_config(
    page_title="ETMAL Calculator",
    layout="wide"
)

st.title("🚢 ETMAL & Invoice Calculator")
st.write(
    "Upload file Excel untuk menghitung "
    "**ETMAL Charge** dan **Invoice (USD)** secara otomatis."
)

# ==============================
# UPLOAD FILE
# ==============================
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
        # Ambil sheet pertama
        xls = pd.ExcelFile(uploaded_file)
        sheet_aktif = xls.sheet_names[0]
        st.info(f"📄 Sheet digunakan: **{sheet_aktif}**")

        df = pd.read_excel(xls, sheet_name=sheet_aktif)

        # ==============================
        # KONVERSI JAM (DP) JADI ANGKA
        # ==============================
        jam_berthing = pd.to_numeric(
            df.iloc[:, 119],
            errors="coerce"
        ).fillna(0)

        # ==============================
        # MAPPING DATA OUTPUT
        # ==============================
        hasil = pd.DataFrame({
            "Name of Vessel": df.iloc[:, 4],
            "Voyage": df.iloc[:, 5],
            "Berth": df.iloc[:, 21],
            "Service": df.iloc[:, 10],
            "ATB": df.iloc[:, 26],
            "ATD": df.iloc[:, 118],
            "No. of Moves": df.iloc[:, 122],
            "TEUS": df.iloc[:, 109],
            "BSH": df.iloc[:, 133],
            "CD": df.iloc[:, 135],
            "GRT": pd.to_numeric(df.iloc[:, 20], errors="coerce").fillna(0),
            "Current Berthing Hours": jam_berthing
        })

        # ==============================
        # PERHITUNGAN
        # ==============================
        hasil["Current Berthing Minutes"] = hasil["Current Berthing Hours"] * 60
        hasil["ETMAL Charge"] = hasil["Current Berthing Hours"].apply(hitung_etmal)
        hasil["Invoice (USD)"] = (
            hasil["GRT"] * 0.131 * hasil["ETMAL Charge"]
        ).round(2)

        # ==============================
        # OUTPUT
        # ==============================
        st.subheader("📊 Hasil Perhitungan ETMAL")
        st.dataframe(hasil, use_container_width=True)

        st.download_button(
            "⬇️ Download Hasil Excel",
            data=hasil.to_excel(index=False),
            file_name="HASIL_ETMAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Terjadi kesalahan saat membaca file.")
        st.code(str(e))
