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
#         "G9999994": "Last Gear Oil Changed",
#         "P9999999": "Last Steering Oil Changed",
#         "W9999999": "Last Steering Oil Changed"
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

#         # For each oil code
#         for code, service in oil_part_mapping.items():
#             sub = group[group["Labour value/part code"] == code].sort_values("Document Date", ascending=False)
#             if sub.empty:
#                 # If we haven’t added this service yet, initialize as N/A
#                 if service not in result:
#                     result[service] = "N/A"
#                 continue

#             latest = sub.iloc[0]
#             km_reading = latest.get("KM/HR Reading", "N/A")
#             entry_text = f"{latest['Document Date'].strftime('%d.%m.%Y')} ({latest['Quantity']} L – {km_reading} KM)"

#             # If we already have one entry (like Steering Oil from both codes), append both
#             if service in result and result[service] != "N/A":
#                 result[service] += f" | {entry_text}"
#             else:
#                 result[service] = entry_text

#             # Handle multiple small-quantity top-ups if needed
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

# WITH CLUTCH OIL--------------------------------------------
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
#         "G9999994": "Last Gear Oil Changed",
#         "P9999999": "Last Steering Oil Changed",
#         "W9999999": "Last Steering Oil Changed",
#         "U9999995": "Last Clutch Oil Changed"
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

#         # For each oil code
#         for code, service in oil_part_mapping.items():
#             sub = group[group["Labour value/part code"] == code].sort_values("Document Date", ascending=False)
#             if sub.empty:
#                 if service not in result:
#                     result[service] = "N/A"
#                 continue

#             latest = sub.iloc[0]
#             km_reading = latest.get("KM/HR Reading", "N/A")
#             entry_text = f"{latest['Document Date'].strftime('%d.%m.%Y')} ({latest['Quantity']} L – {km_reading} KM)"

#             # Combine entries if same service occurs for multiple codes
#             if service in result and result[service] != "N/A":
#                 result[service] += f" | {entry_text}"
#             else:
#                 result[service] = entry_text

#             # Handle multiple small-quantity top-ups if needed
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


# ITERATION 4-------------------------------------- NEXT DUE DATES: 
# import streamlit as st
# import pandas as pd
# import re
# from io import BytesIO
# from dateutil.relativedelta import relativedelta

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

# # Vehicle Type Selection
# vehicle_type = st.radio("Select Vehicle Type", ["Haulage/Tractor", "Bus"])

# # File Uploader
# uploaded_file = st.file_uploader("Upload the Service History Excel File (.xlsx)", type=["xlsx"])

# if uploaded_file:
#     # Read and prepare data
#     df = pd.read_excel(uploaded_file)
#     df["Document Date"] = pd.to_datetime(df["Document Date"], errors="coerce")

#     # --- OIL PART MAPPING ---
#     oil_part_mapping = {
#         "EN699991": "Engine Oil",
#         "GB699991": "Crown Oil",
#         "G9999994": "Gear Oil",
#         "P9999999": "Steering Oil",
#         "W9999999": "Steering Oil",
#         "U9999995": "Clutch Oil"
#     }

#     # --- FILTER MAPPING ---
#     filter_mapping = {
#         "Air Filter": "Air Filter",
#         "Fuel Filter": "Fuel Filter",
#         "Oil Filter": "Oil Filter",
#         "Adblue Tank Filter": "Adblue Tank Filter",
#     }

#     # --- SERVICE INTERVALS BASED ON VEHICLE TYPE ---
#     service_intervals = {
#         "Haulage/Tractor": {
#             "Engine Oil": (80000, 18),
#             "Engine Coolant": (320000, 36),
#             "Gear Oil": (160000, 18),
#             "Steering Oil": (160000, 24),
#             "Crown Oil": (200000, 24),
#             "Clutch Oil": (120000, 12),
#             "Fuel Filter": (80000, 12),
#             "Air Filter": (80000, 12)
#         },
#         "Bus": {
#             "Engine Oil": (80000, 18),
#             "Engine Coolant": (320000, 36),
#             "Gear Oil": (120000, 18),
#             "Steering Oil": (160000, 24),
#             "Crown Oil": (200000, 24),
#             "Clutch Oil": (120000, 12),
#             "Fuel Filter": (80000, 12),
#             "Air Filter": (80000, 12)
#         }
#     }

#     intervals = service_intervals[vehicle_type]

#     # --- FUNCTION TO FETCH OIL ENTRIES ---
#     def get_oil_entries(group):
#         result = {}

#         for code, service in oil_part_mapping.items():
#             sub = group[group["Labour value/part code"] == code].sort_values("Document Date", ascending=False)
#             if sub.empty:
#                 if f"Last {service} Changed" not in result:
#                     result[f"Last {service} Changed"] = "N/A"
#                     result[f"Next {service} Due"] = "N/A"
#                 continue

#             latest = sub.iloc[0]
#             km_reading = latest.get("KM/HR Reading", "N/A")

#             # --- LAST SERVICE INFO ---
#             last_date = latest["Document Date"]
#             result[f"Last {service} Changed"] = f"{last_date.strftime('%d.%m.%Y')} ({latest['Quantity']} L – {km_reading} KM)"

#             # --- NEXT DUE CALCULATION ---
#             if service in intervals and pd.notna(last_date) and pd.notna(km_reading):
#                 km_interval, month_interval = intervals[service]
#                 try:
#                     next_km = int(km_reading) + km_interval
#                 except Exception:
#                     next_km = "N/A"
#                 next_date = last_date + relativedelta(months=month_interval)
#                 result[f"Next {service} Due"] = f"{next_date.strftime('%d.%m.%Y')} ({next_km} KM)"
#             else:
#                 result[f"Next {service} Due"] = "N/A"

#         return pd.Series(result)

#     # --- FUNCTION TO FETCH FILTER ENTRIES ---
#     def get_filter_entries(group):
#         result = {}
#         for key, colname in filter_mapping.items():
#             mask = group["Labour Value/Part description"].str.contains(key, case=False, na=False)
#             mask &= ~group["Labour Value/Part description"].str.contains(r"R\s*&\s*R|R\s*and\s*R", flags=re.I, na=False)
#             sub = group[mask].sort_values("Document Date", ascending=False)
#             if sub.empty:
#                 result[f"Last {colname} Changed"] = "N/A"
#                 result[f"Next {colname} Due"] = "N/A"
#             else:
#                 latest = sub.iloc[0]
#                 km_reading = latest.get("KM/HR Reading", "N/A")
#                 last_date = latest["Document Date"]

#                 result[f"Last {colname} Changed"] = (
#                     f"{last_date.strftime('%d.%m.%Y')} "
#                     f"({latest['Labour Value/Part description']}) – {km_reading} KM"
#                 )

#                 # Next due
#                 if colname in intervals and pd.notna(last_date) and pd.notna(km_reading):
#                     km_interval, month_interval = intervals[colname]
#                     try:
#                         next_km = int(km_reading) + km_interval
#                     except Exception:
#                         next_km = "N/A"
#                     next_date = last_date + relativedelta(months=month_interval)
#                     result[f"Next {colname} Due"] = f"{next_date.strftime('%d.%m.%Y')} ({next_km} KM)"
#                 else:
#                     result[f"Next {colname} Due"] = "N/A"
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
#         temp_df = temp_df[temp_df["Field"] != "Registration number"]
#         temp_df = pd.concat([
#             pd.DataFrame({"Field": ["Registration number"], "Value": [reg_no]}),
#             temp_df
#         ], ignore_index=True)
#         vertical_dataframes.append(temp_df)

#     combined_df = pd.concat(vertical_dataframes, ignore_index=True)

#     # --- EXPORT TO EXCEL ---
#     output = BytesIO()
#     with pd.ExcelWriter(output, engine="openpyxl") as writer:
#         combined_df.to_excel(writer, index=False)
#     output.seek(0)

#     # --- DISPLAY SUCCESS MESSAGE & DOWNLOAD BUTTON ---
#     st.success("✅ Report generated successfully!")
#     st.download_button("📥 Download Excel Report", output, file_name="service_report_with_next_due.xlsx")

# --ITERATION 5-----------------------------------Next Due Dates, with new format: 
# app.py
# import streamlit as st
# import pandas as pd
# import re
# from io import BytesIO
# from dateutil.relativedelta import relativedelta

# st.set_page_config(page_title="YASH MOTORS 🧾 Service Report Generator", layout="centered")

# st.markdown(
#     """
#     <div style="text-align:center;">
#       <h1 style="color:#2E86C1">YASH MOTORS</h1>
#       <h3 style="color:#555">🧾 Service Report Generator</h3>
#     </div>
#     """,
#     unsafe_allow_html=True,
# )

# vehicle_type = st.radio("Select vehicle type", ["Haulage/Tractor", "Bus"])

# uploaded_file = st.file_uploader("Upload the Service History Excel File (.xlsx)", type=["xlsx", "xls"])

# def find_col(df_cols, keywords):
#     """
#     Find the first column name in df_cols that contains any of the keywords (case-insensitive).
#     Returns None if not found.
#     """
#     for k in keywords:
#         for c in df_cols:
#             if k.lower() in c.lower():
#                 return c
#     return None

# def parse_km(value):
#     """Try to extract an integer KM reading from value (handles strings like '52,519', '52519 KM', numeric, etc.)."""
#     if pd.isna(value):
#         return None
#     if isinstance(value, (int, float)) and not pd.isna(value):
#         try:
#             return int(value)
#         except:
#             return None
#     s = str(value)
#     # Common patterns: digits and commas; remove non-digit
#     m = re.search(r"(\d[\d,]*)", s)
#     if m:
#         digits = m.group(1).replace(",", "")
#         try:
#             return int(digits)
#         except:
#             return None
#     return None

# def add_months(dt, months):
#     return dt + relativedelta(months=months)

# # Service intervals (KM, months) as requested
# SERVICE_INTERVALS = {
#     "Haulage/Tractor": {
#         "Engine Oil": (80000, 18),
#         "Engine Coolant": (320000, 36),
#         "Gear Oil": (160000, 18),
#         "Steering Oil": (160000, 24),
#         "Crown Oil": (200000, 24),
#         "Clutch Oil": (120000, 12),
#         "Fuel Filter": (80000, 12),
#         "Air Filter": (80000, 12),
#     },
#     "Bus": {
#         "Engine Oil": (80000, 18),
#         "Engine Coolant": (320000, 36),
#         "Gear Oil": (120000, 18),
#         "Steering Oil": (160000, 24),
#         "Crown Oil": (200000, 24),
#         "Clutch Oil": (120000, 12),
#         "Fuel Filter": (80000, 12),
#         "Air Filter": (80000, 12),
#     }
# }

# # Part codes mapping (labour value/part code)
# PART_CODE_TO_SERVICE = {
#     "EN699991": "Engine Oil",
#     "GB699991": "Crown Oil",
#     "G9999994": "Gear Oil",
#     "P9999999": "Steering Oil",
#     "W9999999": "Steering Oil",
#     "U9999995": "Clutch Oil"
# }

# # Filter textual mapping (we search description)
# FILTER_KEYWORDS = {
#     "Air Filter": "Air Filter",
#     "Fuel Filter": "Fuel Filter",
#     "Oil Filter": "Oil Filter",
#     "Adblue Tank Filter": "Adblue Tank Filter",
# }

# if uploaded_file:
#     try:
#         df_raw = pd.read_excel(uploaded_file, dtype=object)
#     except Exception as e:
#         st.error(f"Failed to read Excel file: {e}")
#         st.stop()

#     # Normalize column names: strip spaces but keep original names for writing
#     original_cols = list(df_raw.columns)
#     cols_lower = [c.strip().lower() for c in original_cols]

#     # Detect important columns with multiple candidate names
#     # You may extend these candidate lists if your sheet uses different header texts
#     doc_date_col = find_col(original_cols, ["Document Date", "Job Card Date", "Date", "Service Date", "Inward Date"])
#     reg_col = find_col(original_cols, ["Registration number", "regis", "registration", "registration no", "reg. no", "vehicle reg"])
#     part_code_col = find_col(original_cols, ["Labour value/part code", "part code", "item number", "labour value/part code", "item"])
#     part_desc_col = find_col(original_cols, ["Labour Value/Part description", "part description", "labour value/part description", "description", "item description"])
#     km_col = find_col(original_cols, ["KM/HR Reading", "km/hr reading", "km reading", "odometer", "odometer reading", "km/hrs", "secondary cumulative reading"])
#     qty_col = find_col(original_cols, ["Quantity", "Qty", "QTY", "Quantity "])

#     # If any critical column not found, show a helpful message listing detected column names
#     missing = []
#     if doc_date_col is None:
#         missing.append("Document Date (e.g. 'Document Date', 'Job Card Date')")
#     if reg_col is None:
#         missing.append("Registration number (vehicle reg)")
#     if part_code_col is None:
#         missing.append("Labour value/part code / Item Number (part code)")
#     if part_desc_col is None:
#         missing.append("Labour Value/Part description / description")
#     if km_col is None:
#         missing.append("KM/HR Reading / Odometer")

#     if missing:
#         st.error("Could not detect required columns in your Excel sheet. Missing: " + "; ".join(missing))
#         st.info("Detected columns: " + ", ".join(original_cols[:20]) + ("..." if len(original_cols) > 20 else ""))
#         st.stop()

#     # Work on a copy with consistent column names
#     df = df_raw.copy()
#     # Ensure doc_date_col is parsed to datetime
#     df[doc_date_col] = pd.to_datetime(df[doc_date_col], errors="coerce")

#     # We'll group by registration number
#     groups = df.groupby(df[reg_col].astype(str).str.strip())

#     intervals = SERVICE_INTERVALS[vehicle_type]

#     # Prepare an in-memory Excel
#     excel_bytes = BytesIO()

#     # We'll write one sheet per registration number (easier to view individual vehicle)
#     with pd.ExcelWriter(excel_bytes, engine="openpyxl") as writer:
#         for reg_value, grp in groups:
#             # Build output_data for this vehicle
#             output_data = {}

#             # 1) Oils & special parts via part code
#             for code, service in PART_CODE_TO_SERVICE.items():
#                 mask = grp[part_code_col].astype(str).str.strip().fillna("") == str(code)
#                 if mask.any():
#                     sub = grp[mask].sort_values(by=doc_date_col, ascending=False)
#                     latest = sub.iloc[0]
#                     last_date = latest[doc_date_col]
#                     qty = latest.get(qty_col, "")
#                     km_reading_val = parse_km(latest.get(km_col, None))
#                     # Format last string similar to your examples
#                     qty_str = f"{qty}" if (qty is not None and str(qty).strip() != "nan") else ""
#                     if pd.notna(last_date):
#                         last_date_str = last_date.strftime("%d.%m.%Y")
#                     else:
#                         last_date_str = "N/A"
#                     last_display = f"{last_date_str}"
#                     if qty_str:
#                         last_display += f" ({qty_str} L"
#                         last_display += f" – {km_reading_val} KM)" if km_reading_val is not None else ")"
#                     else:
#                         last_display += f" ({km_reading_val} KM)" if km_reading_val is not None else ""

#                     output_data[f"Last {service} Changed"] = last_display

#                     # compute next due if interval exists
#                     if service in intervals and pd.notna(last_date) and km_reading_val is not None:
#                         km_interval, month_interval = intervals[service]
#                         next_km = km_reading_val + km_interval
#                         next_date = add_months(last_date, month_interval)
#                         output_data[f"Next {service} Due"] = f"{next_date.strftime('%d.%m.%Y')} ({next_km} KM)"
#                     else:
#                         output_data[f"Next {service} Due"] = ""
#                 else:
#                     output_data[f"Last {service} Changed"] = "N/A"
#                     output_data[f"Next {service} Due"] = ""

#             # 2) Filters by description matching (ignore R&R)
#             for fk, label in FILTER_KEYWORDS.items():
#                 # search in description column for the keyword
#                 desc_series = grp[part_desc_col].astype(str).fillna("")
#                 mask = desc_series.str.contains(label, case=False, na=False)
#                 # exclude R&R
#                 mask &= ~desc_series.str.contains(r"R\s*&\s*R|R\s*and\s*R", case=False, na=False)
#                 if mask.any():
#                     sub = grp[mask].sort_values(by=doc_date_col, ascending=False)
#                     latest = sub.iloc[0]
#                     last_date = latest[doc_date_col]
#                     desc_text = latest.get(part_desc_col, "")
#                     km_reading_val = parse_km(latest.get(km_col, None))
#                     if pd.notna(last_date):
#                         last_date_str = last_date.strftime("%d.%m.%Y")
#                     else:
#                         last_date_str = "N/A"
#                     last_display = f"{last_date_str} ({desc_text})"
#                     if km_reading_val is not None:
#                         last_display += f" – {km_reading_val} KM"

#                     output_data[f"Last {label} Changed"] = last_display

#                     # next due
#                     if label in intervals and pd.notna(last_date) and km_reading_val is not None:
#                         km_interval, month_interval = intervals[label]
#                         next_km = km_reading_val + km_interval
#                         next_date = add_months(last_date, month_interval)
#                         output_data[f"Next {label} Due"] = f"{next_date.strftime('%d.%m.%Y')} ({next_km} KM)"
#                     else:
#                         output_data[f"Next {label} Due"] = ""
#                 else:
#                     output_data[f"Last {label} Changed"] = "N/A"
#                     output_data[f"Next {label} Due"] = ""

#             # Determine list/pairs in the order you want in sheet
#             pairs = [
#                 ("Last Engine Oil Changed", "Next Engine Oil Due"),
#                 ("Last Crown Oil Changed", "Next Crown Oil Due"),
#                 ("Last Gear Oil Changed", "Next Gear Oil Due"),
#                 ("Last Steering Oil Changed", "Next Steering Oil Due"),
#                 ("Last Clutch Oil Changed", "Next Clutch Oil Due"),
#                 ("Last Air Filter Changed", "Next Air Filter Due"),
#                 ("Last Fuel Filter Changed", "Next Fuel Filter Due"),
#                 ("Last Oil Filter Changed", "Next Oil Filter Due"),
#                 ("Last Adblue Tank Filter Changed", "Next Adblue Tank Filter Due"),
#             ]

#             final_rows = []
#             for last_key, next_key in pairs:
#                 last_val = output_data.get(last_key, "N/A")
#                 next_val = output_data.get(next_key, "")
#                 final_rows.append({"Service Item": last_key, str(reg_value): last_val, "Next Due Date": next_val})

#             df_final = pd.DataFrame(final_rows)

#             # Write one sheet per registration number (sheet name limited to 31 chars)
#             safe_name = str(reg_value)[:31] if reg_value else "Unknown"
#             df_final.to_excel(writer, index=False, sheet_name=safe_name)

#             # Optional styling via openpyxl (attempt)
#             try:
#                 from openpyxl import load_workbook
#                 from openpyxl.styles import Font, Alignment
#                 writer.book = writer.book  # already present
#                 ws = writer.sheets[safe_name]
#                 # header bold + center
#                 for cell in ws[1]:
#                     cell.font = Font(bold=True)
#                     cell.alignment = Alignment(horizontal="center", vertical="center")
#                 # wrap text for columns
#                 for col in ws.columns:
#                     for cell in col:
#                         cell.alignment = Alignment(wrap_text=True, vertical="top")
#                 # set reasonable column widths
#                 ws.column_dimensions['A'].width = 35
#                 ws.column_dimensions['B'].width = 50
#                 ws.column_dimensions['C'].width = 40
#             except Exception:
#                 # If openpyxl styling fails, it's not fatal; file still written
#                 pass

#     excel_bytes.seek(0)

#     st.success("✅ Report generated")
#     st.download_button("📥 Download Service Report (one sheet per vehicle)", data=excel_bytes.getvalue(),
#                        file_name="Service_Report_per_vehicle.xlsx",
#                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# FINAL ITERATION-------------------------- NOW INCLUDING TIPPERS

import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import io

# -----------------------
# CONFIGURATION SECTION
# -----------------------

# Service schedule details
SERVICE_SCHEDULES = {
    "Haulage/Tractor": {
        "Engine Oil": (80000, 18),
        "Engine Coolant": (320000, 36),
        "Gear Box Oil": (160000, 18),
        "Power Steering Oil": (160000, 24),
        "Crown Oil": (200000, 24),
        "Clutch Oil": (120000, 12),
        "Fuel Filter": (80000, 12),
        "Air Filter": (80000, 12),
    },
    "Bus": {
        "Engine Oil": (80000, 18),
        "Engine Coolant": (320000, 36),
        "Gear Box Oil": (120000, 18),
        "Power Steering Oil": (160000, 24),
        "Crown Oil": (200000, 24),
        "Clutch Oil": (120000, 12),
        "Fuel Filter": (80000, 12),
        "Air Filter": (80000, 12),
    },
    "Tipper": {
        "Engine Oil": (1000, 18),
        "Engine Coolant": (5000, 36),
        "Gear Box Oil": (3000, 18),
        "Power Steering Oil": (4000, 24),
        "Crown Oil": (2000, 24),
        "Clutch Oil": (2000, 12),
        "Fuel Filter": (1000, 6),
        "Air Filter": (1000, 6),
    }
}

# Mapping of part numbers to items
ITEM_CODES = {
    "Engine Oil": ["A9999999"],
    "Crown Oil": ["B9999999"],
    "Gear Box Oil": ["C9999999"],
    "Power Steering Oil": ["P9999999", "W9999999"],
    "Clutch Oil": ["U9999995"],
    "Fuel Filter": ["F9999999"],
    "Air Filter": ["E9999999"],
    "Oil Filter": ["O9999999"],
    "Adblue Tank Filter": ["T9999999"],
}

# -----------------------
# HELPER FUNCTIONS
# -----------------------

def extract_latest_entry(df, item_codes):
    """Find the latest service entry for given item code(s)."""
    df_filtered = df[df["Item Number"].astype(str).isin(item_codes)]
    if df_filtered.empty:
        return None, None
    df_filtered["Date"] = pd.to_datetime(df_filtered["Document Date"], errors="coerce")
    latest_row = df_filtered.sort_values("Date", ascending=False).iloc[0]
    return latest_row["Date"], latest_row["KM/HR"]

def calculate_next_due(last_date, last_km_or_hr, schedule_km_hr, schedule_months, is_tipper=False):
    """Calculate the next due date and KM/Hr."""
    if pd.isna(last_date) or pd.isna(last_km_or_hr):
        return "N/A"

    try:
        next_date = last_date + pd.DateOffset(months=schedule_months)
        next_km_hr = last_km_or_hr + schedule_km_hr
        unit = "Hrs" if is_tipper else "KM"
        return f"{next_date.strftime('%d.%m.%Y')} or {int(next_km_hr)} {unit}"
    except Exception:
        return "N/A"

# -----------------------
# STREAMLIT APP
# -----------------------

st.title("YASH MOTORS 🧾 Service Report Generator")

uploaded_file = st.file_uploader("Upload the Service History Excel File (.xlsx)", type=["xlsx"])

vehicle_type = st.selectbox("Select Vehicle Type", ["Haulage/Tractor", "Bus", "Tipper"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        if "Item Number" not in df.columns:
            st.error("Error: 'Item Number' column not found in Excel file.")
        else:
            st.success("File uploaded successfully ✅")

            vehicle_no = st.text_input("Enter Vehicle Number:", "CG04PR2489")
            results = []

            for item, codes in ITEM_CODES.items():
                last_date, last_value = extract_latest_entry(df, codes)
                schedule_km_hr, schedule_months = SERVICE_SCHEDULES[vehicle_type].get(item, (None, None))

                if schedule_km_hr and schedule_months:
                    next_due = calculate_next_due(last_date, last_value, schedule_km_hr, schedule_months, vehicle_type == "Tipper")
                else:
                    next_due = "N/A"

                if pd.notna(last_date):
                    formatted = f"{last_date.strftime('%d.%m.%Y')} – {int(last_value)} {'Hrs' if vehicle_type == 'Tipper' else 'KM'}"
                else:
                    formatted = "N/A"

                results.append({
                    "Service Item": f"Last {item} Changed",
                    vehicle_no: formatted,
                    "Next Due Date": next_due
                })

            result_df = pd.DataFrame(results)

            # Display results
            st.dataframe(result_df, use_container_width=True)

            # Option to download
            excel_output = io.BytesIO()
            with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name="Service Report")
            excel_output.seek(0)

            st.download_button(
                label="📥 Download Excel Report",
                data=excel_output,
                file_name=f"{vehicle_no}_Service_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An error occurred while processing: {str(e)}")
