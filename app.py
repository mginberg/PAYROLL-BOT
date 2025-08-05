import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from fuzzywuzzy import process

st.title("Weekly Payroll Calculator with Closers & Enrollers")

# File uploads
hubspot_file = st.file_uploader("Upload HubSpot Deal Tracker CSV")
closer_hours_file = st.file_uploader("Upload Closer Timesheet CSV")
enroller_hours_file = st.file_uploader("Upload Enroller Timesheet CSV")

if hubspot_file and closer_hours_file and enroller_hours_file:
    # Load and process data
    hubspot_df = pd.read_csv(hubspot_file)
    hubspot_df.columns = hubspot_df.columns.str.strip()
    hubspot_df['CLOSER'] = hubspot_df['CLOSER'].str.strip().str.title()
    hubspot_df['ENROLLER'] = hubspot_df['ENROLLER'].astype(str).str.strip().str.title()
    hubspot_df['DATE'] = pd.to_datetime(hubspot_df['DATE'], errors='coerce')

    closer_df = pd.read_csv(closer_hours_file)
    closer_df.columns = closer_df.columns.str.strip()
    closer_df['Rep'] = closer_df['Rep'].str.strip().str.title()
    def parse_and_round_up(duration_str):
        try:
            h, m, s = map(int, duration_str.split(':'))
            total_hours = h + m/60 + s/3600
            return np.ceil(total_hours)
        except:
            return 0
    closer_df['Man Hours'] = closer_df['Man Hours'].astype(str).apply(parse_and_round_up)

    enroller_df = pd.read_csv(enroller_hours_file)
    enroller_df.columns = enroller_df.columns.str.strip()
    enroller_df['Rep'] = enroller_df['Rep'].str.strip().str.title()
    enroller_df['Man Hours'] = enroller_df['Man Hours'].astype(str).apply(parse_and_round_up)

    # Fuzzy Matching to Align Names
    closer_df['Matched Agent'] = closer_df['Rep'].apply(lambda x: process.extractOne(x, hubspot_df['CLOSER'], score_cutoff=80)[0] if process.extractOne(x, hubspot_df['CLOSER'], score_cutoff=80) else x)
    closer_df = closer_df.rename(columns={'Matched Agent': 'Agent'})
    enroller_df['Matched Agent'] = enroller_df['Rep'].apply(lambda x: process.extractOne(x, hubspot_df['ENROLLER'], score_cutoff=80)[0] if process.extractOne(x, hubspot_df['ENROLLER'], score_cutoff=80) else x)
    enroller_df = enroller_df.rename(columns={'Matched Agent': 'Agent'})

    deal_counts = hubspot_df['CLOSER'].value_counts().reset_index()
    deal_counts.columns = ['Agent', 'Deal Count']

    saturday_deals = hubspot_df[hubspot_df['DATE'] == '2025-08-02'].groupby('CLOSER').size().reset_index(name='Saturday Deals')
    saturday_deals['CLOSER'] = saturday_deals['CLOSER'].str.strip().str.title()

    hubspot_df['Deal Date'] = hubspot_df['DATE'].dt.date
    first_deals = hubspot_df.sort_values(by='DATE').drop_duplicates(subset=['Deal Date'], keep='first')
    first_deal_bonus = first_deals['CLOSER'].value_counts().reset_index()
    first_deal_bonus.columns = ['Agent', 'First Deal Bonus Count']

    unique_closers = hubspot_df['CLOSER'].unique()
    unique_closer_df = pd.DataFrame(unique_closers, columns=['Agent'])

    closers = unique_closer_df.merge(closer_df[['Agent', 'Man Hours']], on='Agent', how='left')
    closers = closers.merge(deal_counts, on='Agent', how='left').fillna({'Deal Count': 0})
    closers = closers.merge(saturday_deals.rename(columns={'CLOSER': 'Agent'}), on='Agent', how='left').fillna({'Saturday Deals': 0})
    closers = closers.merge(first_deal_bonus, on='Agent', how='left').fillna({'First Deal Bonus Count': 0})
    closers['Man Hours'] = closers['Man Hours'].fillna(0)

    def determine_hourly_rate(deals):
        if deals >= 15:
            return 22
        elif deals >= 12:
            return 20
        elif deals >= 8:
            return 18
        elif deals >= 4:
            return 15
        else:
            return 13

    def hours_bonus(hours):
        if hours >= 60:
            return 100
        elif hours >= 50:
            return 75
        elif hours >= 40:
            return 50
        else:
            return 0

    closers['Hourly Rate'] = closers['Deal Count'].apply(determine_hourly_rate)
    closers['Hourly Pay'] = closers['Hourly Rate'] * closers['Man Hours']
    closers['Regular Deals'] = closers['Deal Count'] - closers['Saturday Deals']
    closers['Regular Deals Pay'] = closers['Regular Deals'] * 35
    closers['Saturday Deals Pay'] = closers['Saturday Deals'] * 50
    closers['Hours Bonus'] = closers['Man Hours'].apply(hours_bonus)
    closers['First Deal Bonus'] = closers['First Deal Bonus Count'] * 25
    closers['Manual Bonus'] = 0
    closers['$25 Bonus Count'] = 0
    closers['$50 Bonus Count'] = 0

    enroller_submissions = hubspot_df['ENROLLER'].value_counts().reset_index()
    enroller_submissions.columns = ['Agent', 'Submitted Deals']

    enrollers = enroller_df.merge(enroller_submissions, on='Agent', how='left').fillna({'Submitted Deals': 0})
    enrollers['Hourly Rate'] = 18
    enrollers['Hourly Pay'] = enrollers['Man Hours'] * enrollers['Hourly Rate']
    enrollers['Submitted Deals Pay'] = enrollers['Submitted Deals'] * 5
    enrollers['Regular Deals Pay'] = enrollers['Submitted Deals Pay']
    enrollers['Saturday Deals Pay'] = 0
    enrollers['Hours Bonus'] = 0
    enrollers['First Deal Bonus'] = 0
    enrollers['Manual Bonus'] = 0
    enrollers['$25 Bonus Count'] = 0
    enrollers['$50 Bonus Count'] = 0

    closers_export = closers[['Agent', 'Deal Count', 'Man Hours', 'Hourly Rate', 'Hourly Pay', 'Regular Deals Pay', 'Saturday Deals Pay', 'Hours Bonus', 'First Deal Bonus', 'Manual Bonus', '$25 Bonus Count', '$50 Bonus Count']]
    enrollers_export = enrollers[['Agent', 'Submitted Deals', 'Man Hours', 'Hourly Rate', 'Hourly Pay', 'Regular Deals Pay', 'Saturday Deals Pay', 'Hours Bonus', 'First Deal Bonus', 'Manual Bonus', '$25 Bonus Count', '$50 Bonus Count']]
    enrollers_export = enrollers_export.rename(columns={'Submitted Deals': 'Deal Count'})

    combined_export = pd.concat([closers_export, enrollers_export], ignore_index=True, sort=False).fillna(0)

    # XLSX Export with Detailed Formulas
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Payroll Summary"

    headers = ['Agent', 'Deal Count', 'Man Hours', 'Hourly Rate', 'Hourly Pay', 'Regular Deals Pay', 'Saturday Deals Pay', 'Hours Bonus', 'First Deal Bonus', 'Manual Bonus', '$25 Bonus Count', '$50 Bonus Count', 'Total Pay', 'CPA']
    ws.append(headers)

    for idx, row in combined_export.iterrows():
        row_num = idx + 2
        data = [
            row['Agent'],
            row['Deal Count'],
            row['Man Hours'],
            row['Hourly Rate'],
            row['Hourly Pay'],
            row['Regular Deals Pay'],
            row['Saturday Deals Pay'],
            row['Hours Bonus'],
            row['First Deal Bonus'],
            '',  # Manual Bonus
            '',  # $25 Bonus Count
            '',  # $50 Bonus Count
            f"=E{row_num}+F{row_num}+G{row_num}+H{row_num}+I{row_num}+J{row_num}*25+K{row_num}*50+L{row_num}",  # Total Pay Formula
            f"=M{row_num}/B{row_num}" if row['Deal Count'] > 0 else "0"  # CPA
        ]
        ws.append(data)

    total_rows = len(combined_export) + 2
    ws[f"M{total_rows}"] = "Overall CPA:"
    ws[f"N{total_rows}"] = f"=SUM(M2:M{total_rows-1})/SUM(B2:B{total_rows-1})"

    wb.save(output)

    st.download_button("Download Payroll XLSX", output.getvalue(), "Payroll_Summary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.warning("Please upload all three CSV files to proceed.")
