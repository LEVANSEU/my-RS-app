
import streamlit as st
import pandas as pd
from openpyxl import Workbook
import re

st.title("Excel Generator App")

report_file = st.file_uploader("Upload report.xlsx", type=["xlsx"])
statement_file = st.file_uploader("Upload statement.xlsx", type=["xlsx"])

if report_file and statement_file:
    purchases_df = pd.read_excel(report_file, sheet_name='Grid')
    bank_df = pd.read_excel(statement_file)

    purchases_df['დასახელება'] = purchases_df['გამყიდველი'].astype(str).apply(lambda x: re.sub(r'^\(\d+\)\s*', '', x).strip())
    purchases_df['საიდენტიფიკაციო კოდი'] = purchases_df['გამყიდველი'].apply(lambda x: ''.join(re.findall(r'\d', str(x)))[:11])

    bank_df['P'] = bank_df.iloc[:, 15].astype(str).str.strip()
    bank_df['Amount'] = pd.to_numeric(bank_df.iloc[:, 3], errors='coerce').fillna(0)

    st.success("Files uploaded!")
    st.write("Report Preview", purchases_df.head())
    st.write("Statement Preview", bank_df.head())

    if st.button("Generate Final File"):
        wb = Workbook()
        wb.remove(wb.active)

        ws1 = wb.create_sheet(title="ანგარიშფაქტურები კომპანიით")
        ws1.append(['დასახელება', 'საიდენტიფიკაციო კოდი', 'ანგარიშფაქტურის №', 'ანგარიშფაქტურის თანხა', 'ჩარიცხული თანხა'])

        for company_id, group in purchases_df.groupby('საიდენტიფიკაციო კოდი'):
            company_name = group['დასახელება'].iloc[0]
            start_row = ws1.max_row + 1
            unique_invoices = group.groupby('სერია №')['ღირებულება დღგ და აქციზის ჩათვლით'].sum().reset_index()
            company_invoice_sum = unique_invoices['ღირებულება დღგ და აქციზის ჩათვლით'].sum()
            payment_formula = f"=SUMIF(საბანკოამონაწერი!P:P, B{start_row}, საბანკოამონაწერი!D:D)"

            ws1.append([company_name, company_id, '', company_invoice_sum, payment_formula])
            for _, row in unique_invoices.iterrows():
                ws1.append(['', '', row['სერია №'], row['ღირებულება დღგ და აქციზის ჩათვლით'], ''])

        output_path = '/mnt/data/final_file.xlsx'
        wb.save(output_path)
        st.success(f"✅ Final file generated! [Download here](final_file.xlsx)")
