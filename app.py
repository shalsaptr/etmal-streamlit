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
    "Upload file Excel **DATA** untuk menghitung "
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
    etmal = math.ceil(jam / 6) * 0.25
    return min(etmal, 2)

# ==============================
# PROSES FILE
# ==============================
if uploaded_file:
    try:
        # Baca semua sheet
        xls = pd.ExcelFile(uploaded_file)
        sheet_aktif = xls.sheet_names[0]

        st.info(f"📄 Sheet digunakan: **{sheet_aktif}**")

        df = pd.read_excel(xls, sheet_name=sheet_aktif)

        # ==============================
        # MAPPING KOLOM (BERDASARKAN POSISI)
        # ==============================
        hasil = pd.DataFrame({
            "Name of Vessel": df.iloc[:, 4],          # E
            "Voyage": df.iloc[:, 5],                  # F
            "Berth": df.iloc[:, 21],                  # V
            "Service": df.iloc[:, 10],                # K
            "ATB": df.iloc[:, 26],                    # AA
            "ATD": df.iloc[:, 118],                   # DO
            "No. of Moves": df.iloc[:, 122],          # DS
            "TEUS": df.iloc[:, 109],                  # DF
            "BSH": df.iloc[:, 133],                   # ED
            "CD": df.iloc[:, 135],                    # EF
            "GRT": df.iloc[:, 20],                    # U
            "Current Berthing Hours": df.iloc[:, 119] # DP
        })

        # ==============================
        # PERHITUNGAN
        # ==============================
        hasil["Current Berthing Minutes"] = (
            hasil["Current Berthing Hours"].fillna(0) * 60
        )

        hasil["ETMAL Charge"] = hasil["Current Berthing Hours"].apply(hitung_etmal)

        hasil["Invoice (USD)"] = (
            hasil["GRT"].fillna(0) * 0.131 * hasil["ETMAL Charge"]
        ).round(2)

        # ==============================
        # OUTPUT
        # ==============================
        st.subheader("📊 Hasil Perhitungan ETMAL")
        st.dataframe(hasil, use_container_width=True)

        # ==============================
        # DOWNLOAD
        # ==============================
        st.download_button(
            "⬇️ Download Hasil Excel",
            data=hasil.to_excel(index=False),
            file_name="HASIL_ETMAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Terjadi kesalahan saat membaca file.")
        st.code(str(e))
