import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ETMAL Calculator", layout="wide")
st.title("📊 ETMAL Calculator")

uploaded_file = st.file_uploader(
    "Upload file Excel",
    type=["xlsx", "xls"]
)

if uploaded_file:
    try:
        # =========================
        # PILIH SHEET
        # =========================
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox(
            "📄 Sheet digunakan:",
            xls.sheet_names
        )

        # =========================
        # BACA DATA MULAI BARIS KE-10
        # =========================
        raw_df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name,
            skiprows=9,
            header=None
        )

        # =========================
        # AMBIL KOLOM BERDASARKAN POSISI EXCEL
        # =========================
        df = pd.DataFrame({
            "Name of Vessel": raw_df.iloc[:, 4],      # E
            "Voyage": raw_df.iloc[:, 5],              # F
            "Berth": raw_df.iloc[:, 21],              # V
            "Service": raw_df.iloc[:, 10],            # K
            "ATB": raw_df.iloc[:, 26],                # AA
            "ATD": raw_df.iloc[:, 118],               # DO
            "No. of Moves": raw_df.iloc[:, 122],      # DS
            "TEUS": raw_df.iloc[:, 109],              # DF
            "GRT": raw_df.iloc[:, 20],                # U
            "Current Berthing hours": raw_df.iloc[:, 119],  # DP
        })

        # =========================
        # FORMAT ATB & ATD (INI REVISI UTAMA)
        # =========================
        for col in ["ATB", "ATD"]:
            df[col] = pd.to_datetime(
                df[col],
                errors="coerce"
            ).dt.strftime("%d/%m/%Y %H:%M")

        # =========================
        # KONVERSI KE NUMERIK
        # =========================
        df["Current Berthing hours"] = pd.to_numeric(
            df["Current Berthing hours"],
            errors="coerce"
        ).fillna(0)

        df["GRT"] = pd.to_numeric(
            df["GRT"],
            errors="coerce"
        ).fillna(0)

        # =========================
        # CURRENT BERTHING MINUTES
        # =========================
        df["Current Berthing minutes"] = (
            df["Current Berthing hours"] * 60
        )

        # =========================
        # LOGIKA ETMAL SESUAI EXCEL
        # =========================
        def hitung_etmal(jam):
            if jam <= 6:
                return 0.25
            elif jam <= 12:
                return 0.50
            elif jam <= 18:
                return 0.75
            elif jam <= 24:
                return 1.00
            elif jam <= 30:
                return 1.25
            elif jam <= 36:
                return 1.50
            elif jam <= 42:
                return 1.75
            else:
                return 2.00

        df["Etmal charged"] = df["Current Berthing hours"].apply(hitung_etmal)

        # =========================
        # HITUNG INVOICE
        # =========================
        df["Invoice (USD)"] = (
            df["GRT"] * 0.131 * df["Etmal charged"]
        ).round(2)

        # =========================
        # NOMOR BARIS MULAI DARI 1
        # =========================
        df.insert(0, "No", range(1, len(df) + 1))

        # =========================
        # TAMPILKAN DATA
        # =========================
        st.success("✅ ATB & ATD sudah diformat seragam")
        st.dataframe(df, use_container_width=True)

        # =========================
        # DOWNLOAD EXCEL
        # =========================
        buffer = BytesIO()
        df.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            label="⬇️ Download hasil Excel",
            data=buffer,
            file_name="ETMAL_RESULT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Terjadi kesalahan")
        st.code(str(e))
