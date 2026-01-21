import streamlit as st
import pandas as pd
import math

st.set_page_config(page_title="ETMAL Calculator", layout="wide")
st.title("🚢 ETMAL & Invoice Calculator")

st.write("Upload file **DATA.xlsx**, sistem akan otomatis menghitung ETMAL dan Invoice.")

uploaded_file = st.file_uploader("📤 Upload DATA.xlsx", type=["xlsx"])

def hitung_etmal(jam):
    if pd.isna(jam) or jam <= 0:
        return 0
    return min(math.ceil(jam / 6) * 0.25, 2)

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="DATA")

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
        "GRT": df.iloc[:, 20],
        "Current Berthing Hours": df.iloc[:, 119],
    })

    hasil["Current Berthing Minutes"] = hasil["Current Berthing Hours"] * 60
    hasil["ETMAL Charge"] = hasil["Current Berthing Hours"].apply(hitung_etmal)
    hasil["Invoice (USD)"] = (hasil["GRT"] * 0.131 * hasil["ETMAL Charge"]).round(2)

    st.dataframe(hasil, use_container_width=True)

    st.download_button(
        "⬇️ Download Hasil Excel",
        data=hasil.to_excel(index=False),
        file_name="HASIL_ETMAL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
