import streamlit as st
import pandas as pd

st.set_page_config(page_title="ETMAL Calculator", layout="wide")

st.title("📊 ETMAL Calculator")
st.write("Upload file Excel ETMAL")

# =========================
# UPLOAD FILE
# =========================
uploaded_file = st.file_uploader(
    "Upload file Excel",
    type=["xlsx", "xls"]
)

if uploaded_file:
    try:
        # Ambil daftar sheet
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox(
            "📄 Sheet digunakan:",
            xls.sheet_names
        )

        # Baca sheet
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

        # =========================
        # BERSIHKAN NAMA KOLOM
        # =========================
        df.columns = (
            df.columns
            .astype(str)
            .str.replace('\n', ' ', regex=False)
            .str.replace('"', '', regex=False)
            .str.strip()
        )

        # =========================
        # STANDARKAN NAMA KOLOM
        # =========================
        df.columns = (
            df.columns
            .str.upper()
            .str.replace(' ', '_')
            .str.replace('.', '', regex=False)
            .str.replace('(', '', regex=False)
            .str.replace(')', '', regex=False)
        )

        # =========================
        # URUTAN KOLOM FINAL
        # =========================
        urutan_kolom = [
            "NAME_OF_VESSEL",
            "VOYAGE",
            "SERVICE",
            "BERTH",
            "ATB",
            "ATD",
            "CURRENT_BERTHING_HOURS",
            "NO_OF_MOVES",
            "TEUS",
            "BSH",
            "CD",
            "GRT",
            "ETMAL_CHARGED",
            "INVOICE_USD"
        ]

        # Ambil hanya kolom yang ada
        df = df[[c for c in urutan_kolom if c in df.columns]]

        # =========================
        # RENAME UNTUK TAMPILAN
        # =========================
        rename_display = {
            "NAME_OF_VESSEL": "Name of Vessel",
            "NO_OF_MOVES": "No. of Moves",
            "CURRENT_BERTHING_HOURS": "Current Berthing Hours",
            "ETMAL_CHARGED": "Etmal Charged",
            "INVOICE_USD": "Invoice (USD)"
        }

        df = df.rename(columns=rename_display)

        # =========================
        # TAMPILKAN DATA
        # =========================
        st.success("✅ File berhasil diproses")
        st.dataframe(df, use_container_width=True)

        # =========================
        # DOWNLOAD HASIL
        # =========================
        output = df.to_excel(index=False)
        st.download_button(
            label="⬇️ Download hasil Excel",
            data=output,
            file_name="ETMAL_CLEAN.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Terjadi kesalahan saat membaca file.")
        st.code(str(e))
