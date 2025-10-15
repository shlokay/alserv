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
import re
from io import BytesIO

# Streamlit Page Config
st.set_page_config(page_title="Service Report Generator", page_icon="🧾", layout="centered")

# Header Section
st.markdown(
    """
    <div style="text-align: center; margin-bottom: 20px;">
        <h1 style="color:#2E86C1; margin-bottom: 0;">YASH MOTORS</h1>
        <h3 style="color:#555;">🧾 Service Report Generator</h3>
        <hr style="border: 1px solid #ddd;">
    </div>
    """,
    unsafe_allow_html=True
)

# File Uploader
uploaded_file = st.file_uploader("Upload the Service History Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Read and prepare data
    df = pd.read_excel(uploaded_file)
    df["Document Date"] = pd.to_datetime(df["Document Date"], errors="coerce")

    # --- OIL PART MAPPING ---
    oil_part_mapping = {
        "EN699991": "Last Engine Oil Changed",
        "GB699991": "Last Crown Oil Changed",
        "G9999994": "Last Gear Oil Changed",
        "P9999999": "Last Steering Oil Changed",
        "W9999999": "Last Steering Oil Changed"
    }

    # --- FILTER MAPPING ---
    filter_mapping = {
        "Air Filter": "Air Filter",
        "Fuel Filter": "Fuel Filter",
        "Oil Filter": "Oil Filter",
        "Adblue Tank Filter": "Adblue Tank Filter",
    }

    # --- FUNCTION TO FETCH OIL ENTRIES ---
    def get_oil_entries(group):
        result = {}

        # For each oil code
        for code, service in oil_part_mapping.items():
            sub = group[group["Labour value/part code"] == code].sort_values("Document Date", ascending=False)
            if sub.empty:
                # If we haven’t added this service yet, initialize as N/A
                if service not in result:
                    result[service] = "N/A"
                continue

            latest = sub.iloc[0]
            km_reading = latest.get("KM/HR Reading", "N/A")
            entry_text = f"{latest['Document Date'].strftime('%d.%m.%Y')} ({latest['Quantity']} L – {km_reading} KM)"

            # If we already have one entry (like Steering Oil from both codes), append both
            if service in result and result[service] != "N/A":
                result[service] += f" | {entry_text}"
            else:
                result[service] = entry_text

            # Handle multiple small-quantity top-ups if needed
            if latest["Quantity"] < 10 and len(sub) > 1:
                second = sub.iloc[1]
                km2 = second.get("KM/HR Reading", "N/A")
                result[service] += f" | {second['Document Date'].strftime('%d.%m.%Y')} ({second['Quantity']} L – {km2} KM)"
                if (latest["Document Date"].date() == second["Document Date"].date()) and len(sub) > 2:
                    third = sub.iloc[2]
                    km3 = third.get("KM/HR Reading", "N/A")
                    result[service] += f" | {third['Document Date'].strftime('%d.%m.%Y')} ({third['Quantity']} L – {km3} KM)"

        return pd.Series(result)

    # --- FUNCTION TO FETCH FILTER ENTRIES ---
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
                km_reading = latest.get("KM/HR Reading", "N/A")
                result[colname] = (
                    f"{latest['Document Date'].strftime('%d.%m.%Y')} "
                    f"({latest['Labour Value/Part description']}) – {km_reading} KM"
                )
        return pd.Series(result)

    # --- APPLY FUNCTIONS GROUPWISE ---
    oil_result = df.groupby("Registration number", group_keys=False).apply(get_oil_entries).reset_index()
    filter_result = df.groupby("Registration number", group_keys=False).apply(get_filter_entries).reset_index()

    # --- MERGE RESULTS ---
    final_result = pd.merge(oil_result, filter_result, on="Registration number", how="outer")

    # --- CONVERT TO VERTICAL FORMAT ---
    vertical_dataframes = []
    for _, row in final_result.iterrows():
        reg_no = row["Registration number"]
        temp_df = pd.DataFrame({
            "Field": row.index,
            "Value": row.values
        })
        temp_df.columns = ["Field", "Value"]
        temp_df = temp_df[temp_df["Field"] != "Registration number"]
        temp_df = pd.concat([
            pd.DataFrame({"Field": ["Registration number"], "Value": [reg_no]}),
            temp_df
        ], ignore_index=True)
        vertical_dataframes.append(temp_df)

    # Combine all registration numbers one below the other
    combined_df = pd.concat(vertical_dataframes, ignore_index=True)

    # --- EXPORT TO EXCEL ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        combined_df.to_excel(writer, index=False)
    output.seek(0)

    # --- DISPLAY SUCCESS MESSAGE & DOWNLOAD BUTTON ---
    st.success("✅ Report generated successfully!")
    st.download_button("📥 Download Excel Report", output, file_name="vertical_service_report.xlsx")
