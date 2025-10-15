# import streamlit as st
# import pandas as pd
# import re
# from io import BytesIO

# st.set_page_config(page_title="Service Report Generator", page_icon="🧾", layout="centered")

# st.markdown(
#     """
#     <div style="text-align: center; margin-bottom: 20px;">
#         <h1 style="color:#2E86C1; margin-bottom: 0;">YASH MOTORS</h1>
#         <h3 style="color:#555;">🧾 Service Report Generator</h3>
#         <hr style="border: 1px solid #ddd;">
#     </div>
#     """,
#     unsafe_allow_html=True
# )


# uploaded_file = st.file_uploader("Upload the Service History Excel File (.xlsx)", type=["xlsx"])

# if uploaded_file:
#     df = pd.read_excel(uploaded_file)
#     df["Document Date"] = pd.to_datetime(df["Document Date"], errors="coerce")

#     oil_part_mapping = {
#         "EN699991": "Last Engine Oil Changed",
#         "GB699991": "Last Crown Oil Changed",
#         "G9999994": "Last Gear Oil Changed"
#     }

#     filter_mapping = {
#         "Air Filter": "Air Filter",
#         "Fuel Filter": "Fuel Filter",
#         "Oil Filter": "Oil Filter"
#     }

#     def get_oil_entries(group):
#         result = {}
#         for code, service in oil_part_mapping.items():
#             sub = group[group["Labour value/part code"] == code].sort_values("Document Date", ascending=False)
#             if sub.empty:
#                 result[service] = "N/A"
#                 continue
#             latest = sub.iloc[0]
#             entry_text = f"{latest['Document Date'].strftime('%d.%m.%Y')} ({latest['Quantity']} L)"
#             result[service] = entry_text
#             if latest["Quantity"] < 10 and len(sub) > 1:
#                 second = sub.iloc[1]
#                 result[service] += f" | {second['Document Date'].strftime('%d.%m.%Y')} ({second['Quantity']} L)"
#                 if (latest["Document Date"].date() == second["Document Date"].date()) and len(sub) > 2:
#                     third = sub.iloc[2]
#                     result[service] += f" | {third['Document Date'].strftime('%d.%m.%Y')} ({third['Quantity']} L)"
#         return pd.Series(result)

#     def get_filter_entries(group):
#         result = {}
#         for key, colname in filter_mapping.items():
#             mask = group["Labour Value/Part description"].str.contains(key, case=False, na=False)
#             mask &= ~group["Labour Value/Part description"].str.contains(r"R\s*&\s*R|R\s*and\s*R", flags=re.I, na=False)
#             sub = group[mask].sort_values("Document Date", ascending=False)
#             if sub.empty:
#                 result[colname] = "N/A"
#             else:
#                 latest = sub.iloc[0]
#                 result[colname] = f"{latest['Document Date'].strftime('%d.%m.%Y')} ({latest['Labour Value/Part description']})"
#         return pd.Series(result)

#     oil_result = df.groupby("Registration number", group_keys=False).apply(get_oil_entries).reset_index()
#     filter_result = df.groupby("Registration number", group_keys=False).apply(get_filter_entries).reset_index()
#     final_result = pd.merge(oil_result, filter_result, on="Registration number", how="outer")

#     output = BytesIO()
#     final_result.to_excel(output, index=False)
#     output.seek(0)

#     st.success("✅ Report generated successfully!")
#     st.download_button("📥 Download Excel Report", output, file_name="filtered_service_report.xlsx")


# ITERATION 2--------------------------

# import streamlit as st
# import pandas as pd
# import re
# from io import BytesIO

# # Streamlit Page Config
# st.set_page_config(page_title="Service Report Generator", page_icon="🧾", layout="centered")

# # Header Section
# st.markdown(
#     """
#     <div style="text-align: center; margin-bottom: 20px;">
#         <h1 style="color:#2E86C1; margin-bottom: 0;">YASH MOTORS</h1>
#         <h3 style="color:#555;">🧾 Service Report Generator</h3>
#         <hr style="border: 1px solid #ddd;">
#     </div>
#     """,
#     unsafe_allow_html=True
# )

# # File Uploader
# uploaded_file = st.file_uploader("Upload the Service History Excel File (.xlsx)", type=["xlsx"])

# if uploaded_file:
#     # Read and prepare data
#     df = pd.read_excel(uploaded_file)
#     df["Document Date"] = pd.to_datetime(df["Document Date"], errors="coerce")

#     # --- OIL PART MAPPING ---
#     oil_part_mapping = {
#         "EN699991": "Last Engine Oil Changed",
#         "GB699991": "Last Crown Oil Changed",
#         "G9999994": "Last Gear Oil Changed"
#     }

#     # --- FILTER MAPPING ---
#     filter_mapping = {
#         "Air Filter": "Air Filter",
#         "Fuel Filter": "Fuel Filter",
#         "Oil Filter": "Oil Filter",
#         "Adblue Tank Filter": "Adblue Tank Filter",
#     }

#     # --- FUNCTION TO FETCH OIL ENTRIES ---
#     def get_oil_entries(group):
#         result = {}
#         for code, service in oil_part_mapping.items():
#             sub = group[group["Labour value/part code"] == code].sort_values("Document Date", ascending=False)
#             if sub.empty:
#                 result[service] = "N/A"
#                 continue
#             latest = sub.iloc[0]
#             km_reading = latest.get("KM/HR Reading", "N/A")
#             entry_text = f"{latest['Document Date'].strftime('%d.%m.%Y')} ({latest['Quantity']} L – {km_reading} KM)"
#             result[service] = entry_text

#             if latest["Quantity"] < 10 and len(sub) > 1:
#                 second = sub.iloc[1]
#                 km2 = second.get("KM/HR Reading", "N/A")
#                 result[service] += f" | {second['Document Date'].strftime('%d.%m.%Y')} ({second['Quantity']} L – {km2} KM)"
#                 if (latest["Document Date"].date() == second["Document Date"].date()) and len(sub) > 2:
#                     third = sub.iloc[2]
#                     km3 = third.get("KM/HR Reading", "N/A")
#                     result[service] += f" | {third['Document Date'].strftime('%d.%m.%Y')} ({third['Quantity']} L – {km3} KM)"
#         return pd.Series(result)

#     # --- FUNCTION TO FETCH FILTER ENTRIES ---
#     def get_filter_entries(group):
#         result = {}
#         for key, colname in filter_mapping.items():
#             mask = group["Labour Value/Part description"].str.contains(key, case=False, na=False)
#             mask &= ~group["Labour Value/Part description"].str.contains(r"R\s*&\s*R|R\s*and\s*R", flags=re.I, na=False)
#             sub = group[mask].sort_values("Document Date", ascending=False)
#             if sub.empty:
#                 result[colname] = "N/A"
#             else:
#                 latest = sub.iloc[0]
#                 km_reading = latest.get("KM/HR Reading", "N/A")
#                 result[colname] = (
#                     f"{latest['Document Date'].strftime('%d.%m.%Y')} "
#                     f"({latest['Labour Value/Part description']}) – {km_reading} KM"
#                 )
#         return pd.Series(result)

#     # --- APPLY FUNCTIONS GROUPWISE ---
#     oil_result = df.groupby("Registration number", group_keys=False).apply(get_oil_entries).reset_index()
#     filter_result = df.groupby("Registration number", group_keys=False).apply(get_filter_entries).reset_index()

#     # --- MERGE RESULTS ---
#     final_result = pd.merge(oil_result, filter_result, on="Registration number", how="outer")

#     # --- CONVERT TO VERTICAL FORMAT ---
#     vertical_dataframes = []
#     for _, row in final_result.iterrows():
#         reg_no = row["Registration number"]
#         temp_df = pd.DataFrame({
#             "Field": row.index,
#             "Value": row.values
#         })
#         # Move Registration number as a header-like top row
#         temp_df.columns = ["Field", "Value"]
#         temp_df = temp_df[temp_df["Field"] != "Registration number"]
#         temp_df = pd.concat([
#             pd.DataFrame({"Field": ["Registration number"], "Value": [reg_no]}),
#             temp_df
#         ], ignore_index=True)
#         vertical_dataframes.append(temp_df)

#     # Combine all registration numbers one below the other
#     combined_df = pd.concat(vertical_dataframes, ignore_index=True)

#     # --- EXPORT TO EXCEL ---
#     output = BytesIO()
#     with pd.ExcelWriter(output, engine="openpyxl") as writer:
#         combined_df.to_excel(writer, index=False)
#     output.seek(0)

#     # --- DISPLAY SUCCESS MESSAGE & DOWNLOAD BUTTON ---
#     st.success("✅ Report generated successfully!")
#     st.download_button("📥 Download Excel Report", output, file_name="vertical_service_report.xlsx")


# ITERATION 3----------------------------------------------------------------------------
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="YASH MOTORS 🧾 Service Report Generator")

st.title("YASH MOTORS \n🧾 Service Report Generator")

uploaded_file = st.file_uploader("Upload the Service History Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Read Excel file
    df = pd.read_excel(uploaded_file)

    # Normalize column names
    df.columns = df.columns.str.strip().str.lower()

    # Identify possible column names
    date_col = next((c for c in df.columns if "date" in c), None)
    desc_col = next((c for c in df.columns if "description" in c), None)
    item_col = next((c for c in df.columns if "item" in c), None)
    reg_col = next((c for c in df.columns if "registration" in c), None)
    km_col = next((c for c in df.columns if "km" in c or "odometer" in c), None)

    if not all([date_col, desc_col, item_col, reg_col, km_col]):
        st.error("❌ Could not detect all required columns (Service Date, Description, Item Number, Registration, KM). Please check the Excel headers.")
        st.stop()

    # Mapping for filters
    filter_mapping = {
        "Air Filter": "Air Filter",
        "Fuel Filter": "Fuel Filter",
        "Oil Filter": "Oil Filter",
        "Adblue Tank Filter": "Adblue Tank Filter"
    }

    # Function to get last service entry for oils
    def get_last_service(df, keywords):
        mask = df[desc_col].str.contains("|".join(keywords), case=False, na=False)
        filtered = df[mask]
        if filtered.empty:
            return "N/A"
        latest = filtered.sort_values(by=date_col, ascending=False).iloc[0]
        return f"{pd.to_datetime(latest[date_col]).strftime('%d.%m.%Y')} ({latest[desc_col]} – {latest[km_col]} KM)"

    # Function to get last filter change
    def get_filter_entries(df, filter_name):
        mask = df[desc_col].str.contains(filter_mapping[filter_name], case=False, na=False)
        mask &= ~df[desc_col].str.contains("R&R", case=False, na=False)
        filtered = df[mask]
        if filtered.empty:
            return "N/A"
        latest = filtered.sort_values(by=date_col, ascending=False).iloc[0]
        return f"{pd.to_datetime(latest[date_col]).strftime('%d.%m.%Y')} ({latest[desc_col]}) – {latest[km_col]} KM"

    # Function to get last steering oil change (P9999999 or W9999999)
    def get_last_steering_oil(df):
        mask = df[item_col].astype(str).isin(["P9999999", "W9999999"])
        filtered = df[mask]
        if filtered.empty:
            return "N/A"
        latest = filtered.sort_values(by=date_col, ascending=False).iloc[0]
        return f"{pd.to_datetime(latest[date_col]).strftime('%d.%m.%Y')} ({latest[item_col]} – {latest[km_col]} KM)"

    # Process each vehicle
    results = []

    for reg_no, group in df.groupby(reg_col):
        entry = {"Registration number": reg_no}

        entry["Last Engine Oil Changed"] = get_last_service(group, ["Engine Oil"])
        entry["Last Crown Oil Changed"] = get_last_service(group, ["Crown Oil", "Differential Oil"])
        entry["Last Gear Oil Changed"] = get_last_service(group, ["Gear Oil"])
        entry["Last Steering Oil Changed"] = get_last_steering_oil(group)

        for f in filter_mapping:
            entry[f] = get_filter_entries(group, f)

        results.append(entry)

    result_df = pd.DataFrame(results)

    # Create the “vertical” layout Excel export
    excel_output = BytesIO()
    with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
        for _, row in result_df.iterrows():
            reg_no = row["Registration number"]
            temp_df = pd.DataFrame({
                "Field": list(row.index),
                "Value": list(row.values)
            })
            temp_df.to_excel(writer, index=False, sheet_name=str(reg_no)[:31])  # Sheet name ≤ 31 chars

    st.success("✅ Processing complete!")

    st.download_button(
        label="📥 Download Service Report (Formatted)",
        data=excel_output.getvalue(),
        file_name="Service_Report_YashMotors.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
