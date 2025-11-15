# ---- FINAL ITERATION-------------------------------------------------------------------------------

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

# vehicle_type = st.selectbox("Select vehicle type", ["Haulage/Tractor", "Bus", "Tipper"])

# uploaded_file = st.file_uploader("Upload the Service History Excel File (.xlsx)", type=["xlsx", "xls"])

# # ---------- helper functions ----------
# def find_col(df_cols, keywords):
#     """Return first column name that contains any keyword (case-insensitive)."""
#     for k in keywords:
#         for c in df_cols:
#             if k.lower() in str(c).lower():
#                 return c
#     return None

# def parse_km(value):
#     """Extract integer from common km/hr string formats like '52,519', '52519 KM', numeric, etc."""
#     if pd.isna(value):
#         return None
#     if isinstance(value, (int, float)):
#         try:
#             return int(value)
#         except:
#             return None
#     s = str(value)
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

# # ---------- service intervals ----------
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
#         "Oil Filter": (80000, 12),
#         "DEF Tank Filter": (160000, 24),
#         "APDA Filter": (80000, 12),
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
#         "Oil Filter": (80000, 12),
#         "DEF Tank Filter": (160000, 24),
#         "APDA Filter": (80000, 12),
#     },
#     "Tipper": {
#         "Engine Oil": (1000, 18),
#         "Engine Coolant": (5000, 36),
#         "Gear Oil": (3000, 18),
#         "Steering Oil": (4000, 24),
#         "Crown Oil": (2000, 24),
#         "Clutch Oil": (2000, 12),
#         "Fuel Filter": (1000, 6),
#         "Air Filter": (1000, 6),
#         "Oil Filter": (1000, 6),
#         "DEF Tank Filter": (2000, 12),
#         "APDA Filter": (1000, 6),
#     },
# }

# # ---------- all services and their labour value codes ----------
# SERVICE_CODES = {
#     "Engine Oil": ["EN699991"],
#     "Gear Oil": ["G9999994"],
#     "Crown Oil": ["GB699991"],
#     "Air Filter": ["MB442004", "MB442003", "MB442001", "MB442002", "MB442020", "P5105689", "FG802600", "FG802700", "FG800900", "FG801000"],
#     "Fuel Filter": ["P5105607", "P5105608", "P5105609", "FHJ01400", "FHJ01600", "FHJ02300", "FHJ02400"],
#     "Oil Filter": ["F7A05000", "F7A01500"],
#     "DEF Tank Filter": ["PET00001", "XFZ00200"],
#     "Steering Oil": ["P9999999", "W9999999"],
#     "Clutch Oil": ["U9999995"],
#     "Engine Coolant": ["C9999991"],
#     "APDA Filter": ["PD600968", "PD601147", "PD601037"],
# }

# # ---------- main ----------
# if uploaded_file:
#     try:
#         df_raw = pd.read_excel(uploaded_file, dtype=object)
#     except Exception as e:
#         st.error(f"Failed to read Excel file: {e}")
#         st.stop()

#     original_cols = list(df_raw.columns)

#     # detect necessary columns
#     doc_date_col = find_col(original_cols, ["Document Date", "Job Card Date", "Date", "Service Date", "Inward Date"])
#     reg_col = find_col(original_cols, ["Registration number", "regi", "registration", "reg. no", "vehicle reg", "registration no"])
#     part_code_col = find_col(original_cols, ["Labour value/part code", "part code", "item number", "partno", "labour value/part code"])
#     km_col = find_col(original_cols, ["KM/HR Reading", "km/hr reading", "km reading", "odometer", "odometer reading", "km/hrs", "cumulative reading"])
#     qty_col = find_col(original_cols, ["Quantity", "Qty", "QTY"])

#     missing = []
#     if doc_date_col is None: missing.append("Document Date")
#     if reg_col is None: missing.append("Registration number")
#     if part_code_col is None: missing.append("Labour value/part code / Item Number")
#     if km_col is None: missing.append("KM/HR Reading")

#     if missing:
#         st.error("Could not detect required columns in your Excel sheet. Missing: " + "; ".join(missing))
#         st.info("Detected columns (first 20): " + ", ".join(original_cols[:20]) + ("..." if len(original_cols) > 20 else ""))
#         st.stop()

#     df = df_raw.copy()
#     df[doc_date_col] = pd.to_datetime(df[doc_date_col], errors="coerce")

#     intervals = SERVICE_INTERVALS[vehicle_type]
#     groups = df.groupby(df[reg_col].astype(str).str.strip())

#     excel_bytes = BytesIO()
#     with pd.ExcelWriter(excel_bytes, engine="openpyxl") as writer:
#         for reg_value, grp in groups:
#             output_data = {}
#             for service, codes in SERVICE_CODES.items():
#                 mask = grp[part_code_col].astype(str).isin(codes)
#                 if mask.any():
#                     sub = grp[mask].sort_values(by=doc_date_col, ascending=False)
#                     latest = sub.iloc[0]
#                     last_date = latest[doc_date_col] if pd.notna(latest[doc_date_col]) else None
#                     qty = latest.get(qty_col, "")
#                     km_val = parse_km(latest.get(km_col, None))
#                     part_number = str(latest.get(part_code_col, "")).strip()
#                     unit = "HRS" if vehicle_type == "Tipper" else "KM"

#                     # format the "last changed" display
#                     if last_date is not None:
#                         last_date_str = last_date.strftime("%d.%m.%Y")
#                     else:
#                         last_date_str = "N/A"

#                     details = []
#                     if part_number:
#                         details.append(part_number)
#                     if qty and str(qty).strip().lower() != "nan":
#                         details.append(f"{qty} L")
#                     if km_val is not None:
#                         details.append(f"{km_val} {unit}")

#                     if details:
#                         last_display = f"{last_date_str} ({' – '.join(details)})"
#                     else:
#                         last_display = last_date_str

#                     output_data[f"Last {service} Changed"] = last_display

#                     # compute next due
#                     if service in intervals and last_date is not None and km_val is not None:
#                         km_int, month_int = intervals[service]
#                         next_km_hr = km_val + km_int
#                         next_date = add_months(last_date, month_int)
#                         output_data[f"Next {service} Due"] = f"{next_date.strftime('%d.%m.%Y')} ({next_km_hr} {unit})"
#                     else:
#                         output_data[f"Next {service} Due"] = ""
#                 else:
#                     output_data[f"Last {service} Changed"] = "N/A"
#                     output_data[f"Next {service} Due"] = ""

#             # order of rows
#             pairs = [
#                 ("Engine Oil",),
#                 ("Crown Oil",),
#                 ("Gear Oil",),
#                 ("Steering Oil",),
#                 ("Clutch Oil",),
#                 ("Engine Coolant",),
#                 ("Air Filter",),
#                 ("Fuel Filter",),
#                 ("Oil Filter",),
#                 ("DEF Tank Filter",),
#                 ("APDA Filter",),
#             ]

#             final_rows = []
#             for (service,) in pairs:
#                 final_rows.append({
#                     "Service Item": f"Last {service} Changed",
#                     str(reg_value): output_data.get(f"Last {service} Changed", "N/A"),
#                     "Next Due Date": output_data.get(f"Next {service} Due", ""),
#                 })

#             df_final = pd.DataFrame(final_rows)

#             safe_name = str(reg_value)[:31] if reg_value else "Unknown"
#             df_final.to_excel(writer, index=False, sheet_name=safe_name)

#             try:
#                 from openpyxl.styles import Font, Alignment
#                 ws = writer.sheets[safe_name]
#                 for cell in ws[1]:
#                     cell.font = Font(bold=True)
#                     cell.alignment = Alignment(horizontal="center", vertical="center")
#                 ws.column_dimensions['A'].width = 35
#                 ws.column_dimensions['B'].width = 55
#                 ws.column_dimensions['C'].width = 40
#                 for col_cells in ws.columns:
#                     for c in col_cells:
#                         c.alignment = Alignment(wrap_text=True, vertical="top")
#             except Exception:
#                 pass

#     excel_bytes.seek(0)
#     st.success("✅ Report generated successfully")
#     st.download_button(
#         "📥 Download Service Report (one sheet per vehicle)",
#         data=excel_bytes.getvalue(),
#         file_name="Service_Report_per_vehicle.xlsx",
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     )


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
    "Air Filter": ["MB442004", "MB442003", "MB442001", "MB442002", "MB442020", "P5105689", "FG802600", "FG802700", "FG800900", "FG801000", "P5105688"],
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
                    history_lines = []
                    latest = None
                    for _, row in sub.iterrows():
                        date = row.get(doc_date_col)
                        qty = row.get(qty_col, "")
                        km_val = parse_km(row.get(km_col, None))
                        part_number = str(row.get(part_code_col, "")).strip()
                        unit = "HRS" if vehicle_type == "Tipper" else "KM"
                        if pd.notna(date):
                            date_str = date.strftime("%d.%m.%Y")
                        else:
                            date_str = "N/A"

                        line_parts = [date_str]
                        if qty and str(qty).strip().lower() != "nan":
                            line_parts.append(f"{qty} L")
                        if km_val is not None:
                            line_parts.append(f"{km_val} {unit}")
                        if part_number:
                            line_parts.append(f"({part_number})")

                        history_lines.append(" – ".join(line_parts))

                    # keep latest for next due
                    latest = sub.iloc[0]
                    last_date = latest[doc_date_col]
                    last_km_val = parse_km(latest.get(km_col, None))

                    output_data[f"Last {service} Changed"] = "\n".join(history_lines)

                    if service in intervals and pd.notna(last_date) and last_km_val is not None:
                        km_int, month_int = intervals[service]
                        next_km_hr = last_km_val + km_int
                        next_date = add_months(last_date, month_int)
                        output_data[f"Next {service} Due"] = f"{next_date.strftime('%d.%m.%Y')} ({next_km_hr} {unit})"
                    else:
                        output_data[f"Next {service} Due"] = ""
                else:
                    output_data[f"Last {service} Changed"] = "N/A"
                    output_data[f"Next {service} Due"] = ""

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
                    "Service Item": f"{service} Change History",
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
                ws.column_dimensions['B'].width = 70
                ws.column_dimensions['C'].width = 40
                for col_cells in ws.columns:
                    for c in col_cells:
                        c.alignment = Alignment(wrap_text=True, vertical="top")
            except Exception:
                pass

    excel_bytes.seek(0)
    st.success("✅ Report generated successfully")
    st.download_button(
        "📥 Download Service Report (Complete History per Vehicle)",
        data=excel_bytes.getvalue(),
        file_name="Service_Report_with_History.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
