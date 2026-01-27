import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ETMAL & Delay Analyzer", layout="wide")
st.title("📊 ETMAL & Delay Analyzer (Multi Vessel)")

uploaded_file = st.file_uploader(
    "Upload file Excel",
    type=["xlsx", "xls"]
)

# =============================
# HELPER
# =============================
def is_excluded(action_plan):
    if pd.isna(action_plan):
        return False
    txt = str(action_plan).lower()
    return any(k in txt for k in ["sholat", "solat", "mandi"])


def calculate_non_overlap(df):
    df = df.sort_values("From")
    total = 0
    last_end = None

    for _, r in df.iterrows():
        start, end = r["From"], r["To"]
        if pd.isna(start) or pd.isna(end):
            continue

        if last_end is None or start >= last_end:
            total += (end - start).total_seconds() / 60
            last_end = end
        elif end > last_end:
            total += (end - last_end).total_seconds() / 60
            last_end = end

    return round(total / 60, 2)


# =============================
# MAIN
# =============================
if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_event = st.selectbox("📄 Sheet Detail Event", xls.sheet_names)

        # =============================
        # READ EVENT SHEET
        # =============================
        raw = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_event,
            header=0
        )

        # =============================
        # IDENTIFY VESSEL HEADER
        # =============================
        raw["VESSEL_ID"] = None

        vessel_mask = raw.iloc[:, 0].astype(str).str.contains("EVER|CMA|MSC|MAERSK", na=False)

        current_vessel = None
        for i in range(len(raw)):
            if vessel_mask.iloc[i]:
                current_vessel = raw.iloc[i, 0]
            raw.at[i, "VESSEL_ID"] = current_vessel

        raw = raw[raw["VESSEL_ID"].notna()].copy()

        # =============================
        # CLEAN EVENT DATA
        # =============================
        raw["From"] = pd.to_datetime(raw["From"], errors="coerce")
        raw["To"] = pd.to_datetime(raw["To"], errors="coerce")
        raw["Duration"] = pd.to_numeric(raw["Duration"], errors="coerce").fillna(0)

        raw = raw[raw["Event"].notna()]

        # =============================
        # FILTER EXCLUDED EVENTS
        # =============================
        raw["exclude"] = raw["Action Plan"].apply(is_excluded)
        raw = raw[raw["exclude"] == False]

        # =============================
        # PROCESS PER VESSEL
        # =============================
        summary = []
        detail_valid = []

        for vessel, g in raw.groupby("VESSEL_ID"):
            total_delay = calculate_non_overlap(g)
            summary.append({
                "Vessel": vessel,
                "Total Delay (Hours)": total_delay
            })

            for ev, ev_df in g.groupby("Event"):
                ev_delay = calculate_non_overlap(ev_df)
                summary.append({
                    "Vessel": vessel,
                    "Event": ev,
                    "Delay (Hours)": ev_delay
                })

            detail_valid.append(g)

        df_summary = pd.DataFrame(summary)
        df_detail = pd.concat(detail_valid)

        # =============================
        # DISPLAY
        # =============================
        st.subheader("📋 Delay Summary per Vessel")
        st.dataframe(df_summary, use_container_width=True)

        st.subheader("📄 Valid Delay Detail (Cleaned)")
        st.dataframe(df_detail, use_container_width=True)

        # =============================
        # EXPORT
        # =============================
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df_summary.to_excel(writer, sheet_name="SUMMARY_DELAY", index=False)
            df_detail.to_excel(writer, sheet_name="DETAIL_DELAY", index=False)

        buffer.seek(0)

        st.download_button(
            "⬇️ Download Final Excel",
            buffer,
            file_name="DELAY_FINAL_RESULT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Error")
        st.code(str(e))
