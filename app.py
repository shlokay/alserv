import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Service Report Generator", page_icon="🧾", layout="centered")

st.title("🧾 Filtered Service Report Generator")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df["Document Date"] = pd.to_datetime(df["Document Date"], errors="coerce")

    oil_part_mapping = {
        "EN699991": "Last Engine Oil Changed",
        "GB699991": "Last Crown Oil Changed",
        "G9999994": "Last Gear Oil Changed"
    }

    filter_mapping = {
        "Air Filter": "Air Filter",
        "Fuel Filter": "Fuel Filter",
        "Oil Filter": "Oil Filter"
    }

    def get_oil_entries(group):
        result = {}
        for code, service in oil_part_mapping.items():
            sub = group[group["Labour value/part code"] == code].sort_values("Document Date", ascending=False)
            if sub.empty:
                result[service] = "N/A"
                continue
            latest = sub.iloc[0]
            entry_text = f"{latest['Document Date'].strftime('%d.%m.%Y')} ({latest['Quantity']} L)"
            result[service] = entry_text
            if latest["Quantity"] < 10 and len(sub) > 1:
                second = sub.iloc[1]
                result[service] += f" | {second['Document Date'].strftime('%d.%m.%Y')} ({second['Quantity']} L)"
                if (latest["Document Date"].date() == second["Document Date"].date()) and len(sub) > 2:
                    third = sub.iloc[2]
                    result[service] += f" | {third['Document Date'].strftime('%d.%m.%Y')} ({third['Quantity']} L)"
        return pd.Series(result)

    def get_filter_entries(group):
        result = {}
        for key, colname in filter_mapping.items():
            mask = group["Labour Value/Part description"].str.contains(key, case=False, na=False)
            mask &= ~group["Labour Value/Part description"].str.contains(r"R\s*&\s*R|R\s*and\s*R", flags=re.I, na=False)
            sub = group[mask].sort_values("Document Date", ascending=False)
            if sub.empty:
                result[colname] = "N/A"
            else:
                latest = sub.iloc[0]
                result[colname] = f"{latest['Document Date'].strftime('%d.%m.%Y')} ({latest['Labour Value/Part description']})"
        return pd.Series(result)

    oil_result = df.groupby("Registration number", group_keys=False).apply(get_oil_entries).reset_index()
    filter_result = df.groupby("Registration number", group_keys=False).apply(get_filter_entries).reset_index()
    final_result = pd.merge(oil_result, filter_result, on="Registration number", how="outer")

    output = BytesIO()
    final_result.to_excel(output, index=False)
    output.seek(0)

    st.success("✅ Report generated successfully!")
    st.download_button("📥 Download Excel Report", output, file_name="filtered_service_report.xlsx")
