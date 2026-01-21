import streamlit as st
import pandas as pd
import math

st.set_page_config(page_title="ETMAL Calculator", layout="wide")
st.title("📊 ETMAL Calculator")

uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx", "xls"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("📄 Sheet digunakan:", xls.sheet_names)

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
            "Name of Vessel": raw_df.iloc[:, 4],    # E
            "Voyage": raw_df.iloc[:, 5],             # F
            "Berth": raw_df.iloc[:, 21],             # V
            "Service": raw_df.iloc[:, 10],           # K
            "ATB": raw_df.iloc[:, 26],               # AA
            "ATD": raw_df.iloc[:, 118],              # DO
            "No. of Moves": raw_df.iloc[:, 122],     # DS
            "TEUS": raw_df.iloc[:, 109],             # DF
            "BSH": raw_df.iloc[:, 133],              # ED
            "CD": raw_df.iloc[:, 135],               # EF
            "GRT": raw_df.iloc[:, 20],               # U
            "Current Berthing hours": raw_df.iloc[:, 119],  # DP
        })

        # =========================
        # KONVERSI NUMERIK (WAJIB)
        # =========================
        for col in [
            "Current Berthing hours",
            "BSH",
            "CD",
            "GRT"
        ]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # =========================
        # HITUNG ETMAL
        # =========================
        def hitung_etmal(row):
            if row["Current Berthing hours"] <= row["BSH"]:
                return 0
            if row["CD"] == 0:
                return 0
            return math.ceil(
                (row["Current Berthing hours"] - row["BSH"]) / row["CD"]
            )

        df["Etmal charged"] = df.apply(hitung_etmal, axis=1)

        # =========================
        # HITUNG INVOICE
        # =========================
        df["Invoice (USD)"] = df["GRT"] * 0.131 * df["Etmal charged"]

        # =========================
        # TAMPILKAN HASIL
        # =========================
        st.success("✅ Data berhasil diproses")
        st.dataframe(df, use_container_width=True)

        # =========================
        # DOWNLOAD
        # =========================
        output = df.to_excel(index=False)
        st.download_button(
            "⬇️ Download hasil Excel",
            output,
            "ETMAL_RESULT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Terjadi kesalahan")
        st.code(str(e))
