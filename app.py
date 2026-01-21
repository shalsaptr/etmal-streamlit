import streamlit as st
import pandas as pd
from io import BytesIO

# ===============================
# JUDUL APLIKASI
# ===============================
st.set_page_config(page_title="ETMAL Streamlit", layout="wide")
st.title("ETMAL - Excel Processing App")

# ===============================
# UPLOAD FILE
# ===============================
uploaded_file = st.file_uploader(
    "Upload file Excel",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    try:
        # ===============================
        # BACA FILE EXCEL
        # ===============================
        df = pd.read_excel(uploaded_file)

        st.success("File berhasil dibaca")
        st.write("Preview Data:")
        st.dataframe(df.head())

        # ===============================
        # PROSES DATA
        # ===============================
        if "jam_berthing" in df.columns:
            df["jam_berthing"] = pd.to_numeric(
                df["jam_berthing"],
                errors="coerce"
            )
            st.info("Kolom 'jam_berthing' berhasil dikonversi ke numerik")
        else:
            st.warning("Kolom 'jam_berthing' tidak ditemukan")

        # ===============================
        # SIMPAN KE EXCEL (BUFFER)
        # ===============================
        buffer = BytesIO()
        df.to_excel(buffer, index=False)
        buffer.seek(0)

        # ===============================
        # DOWNLOAD BUTTON
        # ===============================
        st.download_button(
            label="Download Hasil Excel",
            data=buffer,
            file_name="hasil_etmal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("Terjadi kesalahan saat membaca file.")
        st.exception(e)

else:
    st.info("Silakan upload file Excel terlebih dahulu")
