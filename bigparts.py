import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="🔍 Customer Voice Keyword Filter", layout="centered")

st.markdown(
    """
    <div style='text-align: center;'>
        <h1 style='color: #2E86C1;'>YASH MOTORS</h1>
        <h3>🧾 Service History Keyword Extractor</h3>
    </div>
    """,
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader("📤 Upload Service History Excel File (.xlsx)", type=["xlsx"])

keyword = st.text_input("🔑 Enter keyword to search in 'Customer voice' column:")

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ File loaded successfully! Columns found: {', '.join(df.columns)}")

        if "Customer voice" not in df.columns:
            st.error("❌ Column 'Customer voice' not found. Please check your Excel file headers.")
        else:
            if keyword:
                filtered_df = df[df["Customer voice"].astype(str).str.contains(keyword, case=False, na=False)]

                if not filtered_df.empty:
                    st.success(f"✅ Found {len(filtered_df)} matching rows for keyword: '{keyword}'")

                    st.dataframe(filtered_df, use_container_width=True)

                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        filtered_df.to_excel(writer, index=False, sheet_name='Filtered Results')
                    excel_data = output.getvalue()

                    st.download_button(
                        label="📥 Download Filtered Excel File",
                        data=excel_data,
                        file_name=f"filtered_results_{keyword}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning(f"⚠️ No rows found containing the keyword '{keyword}' in 'Customer voice'.")
            else:
                st.info("💡 Please enter a keyword to start filtering.")
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
else:
    st.info("📂 Please upload a service history Excel file to begin.")
