import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import timedelta

# =====================================================
# KONFIGURASI HALAMAN
# =====================================================
st.set_page_config(page_title="ETMAL Calculator", layout="wide")
st.title("📊 ETMAL Calculator & Delay TPS Analysis")

# =====================================================
# FUNGSI HITUNG DURASI TANPA OVERLAP
# =====================================================
def hitung_durasi_tanpa_overlap(df):
    if df.empty:
        return 0

    intervals = sorted(
        zip(df["From"], df["To"]),
        key=lambda x: x[0]
    )

    total = timedelta(0)
    start, end = intervals[0]

    for s, e in intervals[1:]:
        if s <= end:  # overlap
            end = max(end, e)
        else:
            total += (end - start)
            start, end = s, e

    total += (end - start)
    return int(total.total_seconds() / 60)

# =====================================================
# UPLOAD FILE
# =====================================================
uploaded_file = st.file_uploader(
    "Upload file Excel",
    type=["xlsx", "xls"]
)

if uploaded_file:
    try:
        # =====================================================
        # PILIH SHEET RAW DATA
        # =====================================================
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox(
            "📄 Sheet digunakan (RAW DATA):",
            xls.sheet_names
        )

        # =====================================================
        # BACA RAW DATA (MULAI BARIS KE-10)
        # =====================================================
        raw_df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name,
            skiprows=9,
            header=None
        )

        # Abaikan 2 baris terakhir
        raw_df = raw_df.iloc[:-2]

        # =====================================================
        # AMBIL KOLOM SESUAI POSISI
        # =====================================================
        df = pd.DataFrame({
            "Name of Vessel": raw_df.iloc[:, 4],      # E
            "Voyage": raw_df.iloc[:, 5],              # F
            "Berth": raw_df.iloc[:, 21],              # V
            "Service": raw_df.iloc[:, 10],            # K
            "ATB": raw_df.iloc[:, 26],                # AA
            "ATD": raw_df.iloc[:, 118],               # DO
            "No. of Moves": raw_df.iloc[:, 122],      # DS
            "TEUS": raw_df.iloc[:, 109],              # DF
            "BSH": raw_df.iloc[:, 133],               # ED
            "GRT": raw_df.iloc[:, 20],                # U
            "Current Berthing hours": raw_df.iloc[:, 119],  # DP
        })

        # =====================================================
        # FORMAT ATB & ATD
        # =====================================================
        for col in ["ATB", "ATD"]:
            df[col] = pd.to_datetime(
                df[col],
                errors="coerce"
            ).dt.strftime("%d/%m/%Y %H:%M")

        # =====================================================
        # KONVERSI NUMERIK
        # =====================================================
        for col in ["Current Berthing hours", "GRT", "BSH"]:
            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            ).fillna(0)

        # =====================================================
        # FILTER BSH < 40
        # =====================================================
        df = df[df["BSH"] < 40].reset_index(drop=True)

        # =====================================================
        # CURRENT BERTHING MINUTES
        # =====================================================
        df["Current Berthing minutes"] = df["Current Berthing hours"] * 60

        # =====================================================
        # HITUNG ETMAL SESUAI EXCEL
        # =====================================================
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

        # =====================================================
        # HITUNG INVOICE
        # =====================================================
        df["Invoice (USD)"] = (
            df["GRT"] * 0.131 * df["Etmal charged"]
        ).round(2)

        # =====================================================
        # NOMOR BARIS
        # =====================================================
        df.insert(0, "No", range(1, len(df) + 1))

        # =====================================================
        # TAMPILKAN DASHBOARD ETMAL
        # =====================================================
        st.subheader("📦 Dashboard ETMAL")
        st.dataframe(df, use_container_width=True)

        # =====================================================
        # ================== DELAY TPS ========================
        # =====================================================
        st.subheader("⏱️ Analisis Delay TPS")

        detail_delay = pd.read_excel(
            uploaded_file,
            sheet_name="DETAIL EVENT"
        )

        summary_event = pd.read_excel(
            uploaded_file,
            sheet_name="SUMMARY EVENT"
        )

        # Pastikan datetime
        for col in ["From", "To"]:
            detail_delay[col] = pd.to_datetime(
                detail_delay[col],
                errors="coerce"
            )

        alasan_list = ["BDW", "YDC", "OTHR", "YCG", "RTR", "TTC", "RNP"]
        hasil_delay = []

        vessel_valid = df["Name of Vessel"].unique()

        for vessel in vessel_valid:
            for alasan in alasan_list:

                df_event = detail_delay[
                    (detail_delay["Vessel"] == vessel) &
                    (detail_delay["Reason"] == alasan)
                ].copy()

                if df_event.empty:
                    continue

                # OTHR: buang sholat / solat / mandi
                if alasan == "OTHR":
                    df_event = df_event[
                        ~df_event["Action Plan"].str.contains(
                            r"sholat|solat|mandi",
                            case=False,
                            na=False
                        )
                    ]

                durasi_unik = hitung_durasi_tanpa_overlap(df_event)

                if durasi_unik > 0:
                    menit_final = durasi_unik
                else:
                    menit_final = summary_event.loc[
                        summary_event["Reason"] == alasan,
                        "TOTAL EVENT"
                    ].sum()

                hasil_delay.append({
                    "Vessel": vessel,
                    "Alasan": alasan,
                    "Jumlah waktu tidak tumpang tindih (menit)": menit_final
                })

        df_delay = pd.DataFrame(hasil_delay)

        st.dataframe(df_delay, use_container_width=True)

        # =====================================================
        # DOWNLOAD EXCEL (2 SHEET)
        # =====================================================
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="ETMAL", index=False)
            df_delay.to_excel(writer, sheet_name="DELAY_TPS", index=False)

        buffer.seek(0)

        st.download_button(
            label="⬇️ Download hasil Excel",
            data=buffer,
            file_name="ETMAL_DAN_DELAY_TPS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Terjadi kesalahan")
        st.code(str(e))
