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
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="YASH MOTORS 🧾 Service Report Generator", layout="centered")

st.title("YASH MOTORS\n🧾 Service Report Generator")
st.write("Upload your service history Excel file to generate a detailed report.")

uploaded_file = st.file_uploader("Upload the Service History Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # ===== Identify Registration Number =====
        reg_num = None
        if 'Registration number' in df.columns:
            reg_num = df['Registration number'].iloc[0]
        elif 'Regn Number' in df.columns:
            reg_num = df['Regn Number'].iloc[0]
        else:
            reg_num = "N/A"

        # ===== Define item mappings =====
        oil_items = {
            "Engine Oil": ["P0000001", "W0000001"],
            "Crown Oil": ["P0000002", "W0000002"],
            "Gear Oil": ["P0000003", "W0000003"],
            "Steering Oil": ["P9999999", "W9999999"],
            "Clutch Oil": ["U9999995"]
        }

        filter_items = {
            "Air Filter": ["P0000004"],
            "Fuel Filter": ["P0000005"],
            "Oil Filter": ["P0000006"],
            "Adblue Tank Filter": ["P0000007"]
        }

        output_data = {}

        # Helper to find latest date and km
        def find_last_change(item_codes):
            df_match = df[df['Item Number'].isin(item_codes)]
            if df_match.empty:
                return None, None, None
            latest_row = df_match.sort_values(by="Job Card Date", ascending=False).iloc[0]
            date = latest_row["Job Card Date"].strftime("%d.%m.%Y") if pd.notna(latest_row["Job Card Date"]) else "N/A"
            km = latest_row["Odometer Reading"] if pd.notna(latest_row["Odometer Reading"]) else "N/A"
            desc = latest_row["Item Description"] if "Item Description" in df.columns else ""
            return date, km, desc

        # === Calculate due date based on logic ===
        def get_next_due(item_type, last_date, last_km):
            if last_date == "N/A" or last_km == "N/A":
                return None, None

            last_date = pd.to_datetime(last_date, format="%d.%m.%Y", errors='coerce')
            if pd.isna(last_date):
                return None, None

            if item_type == "Engine Oil":
                next_date = last_date + pd.DateOffset(months=18)
                next_km = int(last_km) + 80000
            elif item_type == "Crown Oil":
                next_date = last_date + pd.DateOffset(years=3)
                next_km = int(last_km) + 320000
            elif item_type == "Gear Oil":
                next_date = last_date + pd.DateOffset(months=18)
                next_km = int(last_km) + 160000
            elif item_type == "Steering Oil":
                next_date = last_date + pd.DateOffset(years=3)
                next_km = int(last_km) + 320000
            elif item_type == "Clutch Oil":
                next_date = last_date + pd.DateOffset(years=2)
                next_km = int(last_km) + 160000
            elif item_type == "Air Filter":
                next_date = last_date + pd.DateOffset(years=1)
                next_km = int(last_km) + 80000
            elif item_type == "Fuel Filter":
                next_date = last_date + pd.DateOffset(years=1)
                next_km = int(last_km) + 80000
            elif item_type == "Oil Filter":
                next_date = last_date + pd.DateOffset(years=2)
                next_km = int(last_km) + 160000
            elif item_type == "Adblue Tank Filter":
                next_date = last_date + pd.DateOffset(years=2)
                next_km = int(last_km) + 160000
            else:
                return None, None

            return next_date.strftime("%d.%m.%Y"), next_km

        # ===== Extract Last and Next data =====
        for item, codes in oil_items.items():
            last_date, km, desc = find_last_change(codes)
            if last_date:
                output_data[f"Last {item} Changed"] = f"{last_date} ({desc} – {km} KM)"
                next_date, next_km = get_next_due(item, last_date, km)
                if next_date:
                    output_data[f"Next {item} Oil Due"] = f"{next_date} ({next_km} KM)"
            else:
                output_data[f"Last {item} Changed"] = "N/A"
                output_data[f"Next {item} Oil Due"] = "N/A"

        for item, codes in filter_items.items():
            last_date, km, desc = find_last_change(codes)
            if last_date:
                output_data[f"Last {item} Changed"] = f"{last_date} ({desc}) – {km} KM"
                next_date, next_km = get_next_due(item, last_date, km)
                if next_date:
                    output_data[f"Next {item} Filter Due"] = f"{next_date} ({next_km} KM)"
            else:
                output_data[f"Last {item} Changed"] = "N/A"
                output_data[f"Next {item} Filter Due"] = "N/A"

        # === Final Report ===
        final_rows = []
        registration_number = reg_num if reg_num else "N/A"

        pairs = [
            ("Last Engine Oil Changed", "Next Engine Oil Due"),
            ("Last Crown Oil Changed", "Next Crown Oil Due"),
            ("Last Gear Oil Changed", "Next Gear Oil Due"),
            ("Last Steering Oil Changed", "Next Steering Oil Due"),
            ("Last Clutch Oil Changed", "Next Clutch Oil Due"),
            ("Last Air Filter Changed", "Next Air Filter Due"),
            ("Last Fuel Filter Changed", "Next Fuel Filter Due"),
            ("Last Oil Filter Changed", "Next Oil Filter Due"),
            ("Last Adblue Tank Filter Changed", "Next Adblue Tank Filter Due"),
        ]

        for last_key, next_key in pairs:
            last_val = output_data.get(last_key, "N/A")
            next_val = output_data.get(next_key, "")
            final_rows.append({
                "Service Item": last_key,
                registration_number: last_val,
                "Next Due Date": next_val
            })

        df_final = pd.DataFrame(final_rows)

        # === Export to Excel ===
        excel_output = io.BytesIO()
        with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Service Report')

            workbook = writer.book
            worksheet = writer.sheets['Service Report']

            # Format
            header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9E1F2'})
            cell_format = workbook.add_format({'text_wrap': True, 'valign': 'top', 'border': 1})
            due_format = workbook.add_format({'bold': True, 'font_color': 'green', 'border': 1, 'valign': 'top'})

            worksheet.set_column('A:A', 35, cell_format)
            worksheet.set_column('B:B', 50, cell_format)
            worksheet.set_column('C:C', 40, due_format)
            worksheet.set_row(0, None, header_format)

        st.success("✅ Report generated successfully!")
        st.download_button("📥 Download Service Report", data=excel_output.getvalue(), file_name="Service_Report.xlsx")

    except Exception as e:
        st.error(f"Error: {str(e)}")
