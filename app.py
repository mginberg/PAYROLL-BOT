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

    # Fuzzy match to align names from Hours → Deals
    def fuzzy_match(name, name_list):
        match = process.extractOne(name, name_list, score_cutoff=80)
        return match[0] if match else name

    closer_df['Agent'] = closer_df['Rep'].apply(lambda x: fuzzy_match(x, hubspot_df['CLOSER'].unique()))
    hubspot_df['Matched Agent'] = hubspot_df['CLOSER'].apply(lambda x: fuzzy_match(x, closer_df['Agent'].unique()))

    # Aggregate Deal Counts
    deal_counts = hubspot_df['Matched Agent'].value_counts().reset_index()
    deal_counts.columns = ['Agent', 'Deal Count']

    # Saturday Deals Count
    saturday_deals = hubspot_df[hubspot_df['DATE'] == '2025-08-02']['Matched Agent'].value_counts().reset_index()
    saturday_deals.columns = ['Agent', 'Saturday Deals']

    # First Deal Bonus Count
    hubspot_df['Deal Date'] = hubspot_df['DATE'].dt.date
    first_deals = hubspot_df.sort_values(by='DATE').drop_duplicates(subset=['Deal Date'], keep='first')
    first_deal_bonus = first_deals['Matched Agent'].value_counts().reset_index()
    first_deal_bonus.columns = ['Agent', 'First Deal Bonus Count']

    # Merge all counts and timesheet
    payroll_df = deal_counts.merge(closer_df[['Agent', 'Man Hours']], on='Agent', how='left')
    payroll_df = payroll_df.merge(saturday_deals, on='Agent', how='left').merge(first_deal_bonus, on='Agent', how='left')
    payroll_df = payroll_df.fillna({'Man Hours': 0, 'Saturday Deals': 0, 'First Deal Bonus Count': 0})

    # Calculate Pay Components
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

    payroll_df['Hourly Rate'] = payroll_df['Deal Count'].apply(determine_hourly_rate)
    payroll_df['Hourly Pay'] = payroll_df['Hourly Rate'] * payroll_df['Man Hours']
    payroll_df['Regular Deals'] = payroll_df['Deal Count'] - payroll_df['Saturday Deals']
    payroll_df['Regular Deals Pay'] = payroll_df['Regular Deals'] * 35
    payroll_df['Saturday Deals Pay'] = payroll_df['Saturday Deals'] * 50
    payroll_df['Hours Bonus'] = payroll_df['Man Hours'].apply(hours_bonus)
    payroll_df['First Deal Bonus'] = payroll_df['First Deal Bonus Count'] * 25
    payroll_df['Manual Bonus'] = 0
    payroll_df['$25 Bonus Count'] = 0
    payroll_df['$50 Bonus Count'] = 0

    # XLSX Export
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Payroll Summary"

    headers = ['Agent', 'Deal Count', 'Man Hours', 'Hourly Rate', 'Hourly Pay', 'Regular Deals Pay', 'Saturday Deals Pay', 'Hours Bonus', 'First Deal Bonus', 'Manual Bonus', '$25 Bonus Count', '$50 Bonus Count', 'Total Pay', 'CPA']
    ws.append(headers)

    for idx, row in payroll_df.iterrows():
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
            f"=E{row_num}+F{row_num}+G{row_num}+H{row_num}+I{row_num}+J{row_num}+K{row_num}*25+L{row_num}*50",  # Corrected Formula
            f"=M{row_num}/B{row_num}" if row['Deal Count'] > 0 else "0"  # CPA
        ]
        ws.append(data)

    total_rows = len(payroll_df) + 2
    ws[f"M{total_rows}"] = "Overall CPA:"
    ws[f"N{total_rows}"] = f"=SUM(M2:M{total_rows-1})/SUM(B2:B{total_rows-1})"

    wb.save(output)

    st.download_button("Download Payroll XLSX", output.getvalue(), "Payroll_Summary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.warning("Please upload all three CSV files to proceed.")
