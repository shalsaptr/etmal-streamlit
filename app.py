import streamlit as st
import pandas as pd
import math

# ---------------------------------
# KONFIGURASI HALAMAN
# ---------------------------------
st.set_page_config(
    page_title="ETMAL Calculator",
    layout="wide"
)

st.title("🚢 ETMAL & Invoice Calculator")
st.write(
    "Upload file **DATA.xlsx**, sistem akan otomatis menghitung "
    "**ETMAL Charge** dan **Invoice (USD)**."
)

# ---------------------------------
# UPLOAD FILE
# ---------------------------------
uploaded_file = st.file_uploader(
    "📤 Upload file Excel",
    type=["xlsx"]
)

# ---------------------------------
# FUNGSI HITUNG ETMAL
# ---------------------------------
def hitung_etmal(jam):
    if pd.isna(jam) or jam <= 0:
        return 0
    return min(math.ceil(jam / 6) * 0.25, 2)

# ---------------------------------
# PROSES DATA
# ---------------------------------
if uploaded_file is not None:
    try:
        # Baca sheet pertama (apa pun namanya)
        df = pd.read_excel(uploaded_file)

        # Mapping kolom berdasarkan POSISI kolom Excel
        hasil = pd.DataFrame({
            "Name of Vessel": df.iloc[:, 4],        # E
            "Voyage": df.iloc[:, 5],                # F
            "Berth": df.iloc[:, 21],                # V
            "Service": df.iloc[:, 10],               # K
            "ATB": df.iloc[:, 26],                  # AA
            "ATD": df.iloc[:, 118],                 # DO
            "No. of Moves": df.iloc[:, 122],        # DS
            "TEUS": df.iloc[:, 109],                # DF
            "BSH": df.iloc[:, 133],                 # ED
            "CD": df.iloc[:, 135],                  # EF
            "GRT": df.iloc[:, 20],                  # U
            "Current Berthing Hours": df.iloc[:, 119],  # DP
        })

        # Hitungan tambahan
        hasil["Current Berthing Minutes"] = hasil["Current Berthing Hours"] * 60
        hasil["ETMAL Charge"] = hasil["Current Berthing Hours"].apply(hitung_etmal)
        hasil["Invoice (USD)"] = (
            hasil["GRT"] * 0.131 * hasil["ETMAL Charge"]
        ).round(2)

        # ---------------------------------
        # TAMPILKAN HASIL
        # ---------------------------------
        st.subheader("📊 Hasil Perhitungan ETMAL")
        st.dataframe(hasil, use_container_width=True)

        # ---------------------------------
        # DOWNLOAD HASIL
        # ---------------------------------
        st.download_button(
            "⬇️ Download Hasil Excel",
            data=hasil.to_excel(index=False),
            file_name="HASIL_ETMAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except IndexError:
        st.error(
            "❌ Struktur kolom Excel tidak sesuai.\n\n"
            "Pastikan file yang di-upload adalah **DATA.xlsx** "
            "dengan jumlah kolom lengkap."
        )

    except Exception as e:
        st.error(f"❌ Terjadi kesalahan: {e}")

else:
    st.info("⬆️ Silakan upload file Excel untuk memulai perhitungan.")
