
import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📊 Excel Generator App")

report_file = st.file_uploader("Upload report.xlsx", type=["xlsx"])
statement_file = st.file_uploader("Upload statement.xlsx", type=["xlsx"])

if report_file and statement_file:
    try:
        report_file.seek(0)
        statement_file.seek(0)

        report_df = pd.read_excel(report_file)
        statement_df = pd.read_excel(statement_file)

        st.success("✅ Files uploaded!")
        st.subheader("Report Preview")
        st.dataframe(report_df.head())

        st.subheader("Statement Preview")
        st.dataframe(statement_df.head())

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            report_df.to_excel(writer, sheet_name='Report', index=False)
            statement_df.to_excel(writer, sheet_name='Statement', index=False)

        st.success("✅ Final file generated!")

        st.download_button(
            label="⬇️ Download Final Excel",
            data=output.getvalue(),
            file_name="final_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.warning("Please upload both Excel files.")
