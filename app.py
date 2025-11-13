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

# ---- FINAL ITERATION-------------------------------------------------------------------------------

import streamlit as st
import pandas as pd
import re
from io import BytesIO
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="YASH MOTORS 🧾 Service Report Generator", layout="centered")

st.markdown(
    """
    <div style="text-align:center;">
      <h1 style="color:#2E86C1">YASH MOTORS</h1>
      <h3 style="color:#555">🧾 Service Report Generator</h3>
    </div>
    """,
    unsafe_allow_html=True,
)

vehicle_type = st.selectbox("Select vehicle type", ["Haulage/Tractor", "Bus", "Tipper"])

uploaded_file = st.file_uploader("Upload the Service History Excel File (.xlsx)", type=["xlsx", "xls"])

# ---------- helper functions ----------
def find_col(df_cols, keywords):
    """Return first column name that contains any keyword (case-insensitive)."""
    for k in keywords:
        for c in df_cols:
            if k.lower() in str(c).lower():
                return c
    return None

def parse_km(value):
    """Extract integer from common km/hr string formats like '52,519', '52519 KM', numeric, etc."""
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        try:
            return int(value)
        except:
            return None
    s = str(value)
    m = re.search(r"(\d[\d,]*)", s)
    if m:
        digits = m.group(1).replace(",", "")
        try:
            return int(digits)
        except:
            return None
    return None

def add_months(dt, months):
    return dt + relativedelta(months=months)

# ---------- service intervals ----------
SERVICE_INTERVALS = {
    "Haulage/Tractor": {
        "Engine Oil": (80000, 18),
        "Engine Coolant": (320000, 36),
        "Gear Oil": (160000, 18),
        "Steering Oil": (160000, 24),
        "Crown Oil": (200000, 24),
        "Clutch Oil": (120000, 12),
        "Fuel Filter": (80000, 12),
        "Air Filter": (80000, 12),
        "Oil Filter": (80000, 12),
        "DEF Tank Filter": (160000, 24),
        "APDA Filter": (80000, 12),
    },
    "Bus": {
        "Engine Oil": (80000, 18),
        "Engine Coolant": (320000, 36),
        "Gear Oil": (120000, 18),
        "Steering Oil": (160000, 24),
        "Crown Oil": (200000, 24),
        "Clutch Oil": (120000, 12),
        "Fuel Filter": (80000, 12),
        "Air Filter": (80000, 12),
        "Oil Filter": (80000, 12),
        "DEF Tank Filter": (160000, 24),
        "APDA Filter": (80000, 12),
    },
    "Tipper": {
        "Engine Oil": (1000, 18),
        "Engine Coolant": (5000, 36),
        "Gear Oil": (3000, 18),
        "Steering Oil": (4000, 24),
        "Crown Oil": (2000, 24),
        "Clutch Oil": (2000, 12),
        "Fuel Filter": (1000, 6),
        "Air Filter": (1000, 6),
        "Oil Filter": (1000, 6),
        "DEF Tank Filter": (2000, 12),
        "APDA Filter": (1000, 6),
    },
}

# ---------- all services and their labour value codes ----------
SERVICE_CODES = {
    "Engine Oil": ["EN699991"],
    "Gear Oil": ["G9999994"],
    "Crown Oil": ["GB699991"],
    "Air Filter": ["MB442004", "MB442003", "MB442001", "MB442002", "MB442020"],
    "Fuel Filter": ["P5105607", "P5105608", "P5105609", "FHJ01400", "FHJ01600", "FHJ02300", "FHJ02400"],
    "Oil Filter": ["F7A05000", "F7A01500"],
    "DEF Tank Filter": ["PET00001", "XFZ00200"],
    "Steering Oil": ["P9999999", "W9999999"],
    "Clutch Oil": ["U9999995"],
    "Engine Coolant": ["C9999991"],
    "APDA Filter": ["PD600968", "PD601147", "PD601037"],
}

# ---------- main ----------
if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, dtype=object)
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    original_cols = list(df_raw.columns)

    # detect necessary columns
    doc_date_col = find_col(original_cols, ["Document Date", "Job Card Date", "Date", "Service Date", "Inward Date"])
    reg_col = find_col(original_cols, ["Registration number", "regi", "registration", "reg. no", "vehicle reg", "registration no"])
    part_code_col = find_col(original_cols, ["Labour value/part code", "part code", "item number", "partno", "labour value/part code"])
    km_col = find_col(original_cols, ["KM/HR Reading", "km/hr reading", "km reading", "odometer", "odometer reading", "km/hrs", "cumulative reading"])
    qty_col = find_col(original_cols, ["Quantity", "Qty", "QTY"])

    missing = []
    if doc_date_col is None: missing.append("Document Date")
    if reg_col is None: missing.append("Registration number")
    if part_code_col is None: missing.append("Labour value/part code / Item Number")
    if km_col is None: missing.append("KM/HR Reading")

    if missing:
        st.error("Could not detect required columns in your Excel sheet. Missing: " + "; ".join(missing))
        st.info("Detected columns (first 20): " + ", ".join(original_cols[:20]) + ("..." if len(original_cols) > 20 else ""))
        st.stop()

    df = df_raw.copy()
    df[doc_date_col] = pd.to_datetime(df[doc_date_col], errors="coerce")

    intervals = SERVICE_INTERVALS[vehicle_type]
    groups = df.groupby(df[reg_col].astype(str).str.strip())

    excel_bytes = BytesIO()
    with pd.ExcelWriter(excel_bytes, engine="openpyxl") as writer:
        for reg_value, grp in groups:
            output_data = {}
            for service, codes in SERVICE_CODES.items():
                mask = grp[part_code_col].astype(str).isin(codes)
                if mask.any():
                    sub = grp[mask].sort_values(by=doc_date_col, ascending=False)
                    latest = sub.iloc[0]
                    last_date = latest[doc_date_col] if pd.notna(latest[doc_date_col]) else None
                    qty = latest.get(qty_col, "")
                    km_val = parse_km(latest.get(km_col, None))
                    part_number = str(latest.get(part_code_col, "")).strip()
                    unit = "HRS" if vehicle_type == "Tipper" else "KM"

                    # format the "last changed" display
                    if last_date is not None:
                        last_date_str = last_date.strftime("%d.%m.%Y")
                    else:
                        last_date_str = "N/A"

                    details = []
                    if part_number:
                        details.append(part_number)
                    if qty and str(qty).strip().lower() != "nan":
                        details.append(f"{qty} L")
                    if km_val is not None:
                        details.append(f"{km_val} {unit}")

                    if details:
                        last_display = f"{last_date_str} ({' – '.join(details)})"
                    else:
                        last_display = last_date_str

                    output_data[f"Last {service} Changed"] = last_display

                    # compute next due
                    if service in intervals and last_date is not None and km_val is not None:
                        km_int, month_int = intervals[service]
                        next_km_hr = km_val + km_int
                        next_date = add_months(last_date, month_int)
                        output_data[f"Next {service} Due"] = f"{next_date.strftime('%d.%m.%Y')} ({next_km_hr} {unit})"
                    else:
                        output_data[f"Next {service} Due"] = ""
                else:
                    output_data[f"Last {service} Changed"] = "N/A"
                    output_data[f"Next {service} Due"] = ""

            # order of rows
            pairs = [
                ("Engine Oil",),
                ("Crown Oil",),
                ("Gear Oil",),
                ("Steering Oil",),
                ("Clutch Oil",),
                ("Engine Coolant",),
                ("Air Filter",),
                ("Fuel Filter",),
                ("Oil Filter",),
                ("DEF Tank Filter",),
                ("APDA Filter",),
            ]

            final_rows = []
            for (service,) in pairs:
                final_rows.append({
                    "Service Item": f"Last {service} Changed",
                    str(reg_value): output_data.get(f"Last {service} Changed", "N/A"),
                    "Next Due Date": output_data.get(f"Next {service} Due", ""),
                })

            df_final = pd.DataFrame(final_rows)

            safe_name = str(reg_value)[:31] if reg_value else "Unknown"
            df_final.to_excel(writer, index=False, sheet_name=safe_name)

            try:
                from openpyxl.styles import Font, Alignment
                ws = writer.sheets[safe_name]
                for cell in ws[1]:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                ws.column_dimensions['A'].width = 35
                ws.column_dimensions['B'].width = 55
                ws.column_dimensions['C'].width = 40
                for col_cells in ws.columns:
                    for c in col_cells:
                        c.alignment = Alignment(wrap_text=True, vertical="top")
            except Exception:
                pass

    excel_bytes.seek(0)
    st.success("✅ Report generated successfully")
    st.download_button(
        "📥 Download Service Report (one sheet per vehicle)",
        data=excel_bytes.getvalue(),
        file_name="Service_Report_per_vehicle.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
