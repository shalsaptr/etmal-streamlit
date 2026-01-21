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

        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

        # =========================
        # BERSIHKAN NAMA KOLOM
        # =========================
        df.columns = (
            df.columns.astype(str)
            .str.replace("\n", " ", regex=False)
            .str.replace('"', "", regex=False)
            .str.strip()
            .str.upper()
            .str.replace(" ", "_")
        )

        # =========================
        # STANDAR KOLOM
        # =========================
        kolom_map = {
            "NAME_OF_VESSEL": "NAME_OF_VESSEL",
            "VOYAGE": "VOYAGE",
            "BERTH": "BERTH",
            "SERVICE": "SERVICE",
            "ATB": "ATB",
            "ATD": "ATD",
            "NO._OF_MOVES": "NO_OF_MOVES",
            "NO_OF_MOVES": "NO_OF_MOVES",
            "TEUS": "TEUS",
            "BSH": "BSH",
            "CD": "CD",
            "GRT": "GRT",
            "CURRENT_BERTHING_HOURS": "CURRENT_BERTHING_HOURS"
        }

        df = df.rename(columns={k: v for k, v in kolom_map.items() if k in df.columns})

        # =========================
        # KONVERSI NUMERIK (INI KUNCI FIX ERROR)
        # =========================
        kolom_angka = [
            "CURRENT_BERTHING_HOURS",
            "BSH",
            "CD",
            "GRT"
        ]

        for col in kolom_angka:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # =========================
        # HITUNG ETMAL
        # =========================
        def hitung_etmal(row):
            if row["CURRENT_BERTHING_HOURS"] <= row["BSH"]:
                return 0
            else:
                selisih = row["CURRENT_BERTHING_HOURS"] - row["BSH"]
                if row["CD"] == 0:
                    return 0
                return math.ceil(selisih / row["CD"])

        df["ETMAL_CHARGED"] = df.apply(hitung_etmal, axis=1)

        # =========================
        # HITUNG INVOICE
        # =========================
        df["INVOICE_USD"] = df["GRT"] * 0.131 * df["ETMAL_CHARGED"]

        # =========================
        # TAMPILAN AKHIR
        # =========================
        tampil = [
            "NAME_OF_VESSEL",
            "VOYAGE",
            "BERTH",
            "SERVICE",
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

        tampil = [c for c in tampil if c in df.columns]
        df = df[tampil]

        df = df.rename(columns={
            "NAME_OF_VESSEL": "Name of Vessel",
            "CURRENT_BERTHING_HOURS": "Current Berthing Hours",
            "NO_OF_MOVES": "No. of Moves",
            "ETMAL_CHARGED": "Etmal Charged",
            "INVOICE_USD": "Invoice (USD)"
        })

        st.success("✅ Berhasil dihitung")
        st.dataframe(df, use_container_width=True)

        # DOWNLOAD
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
