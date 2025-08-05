import streamlit as st
import pandas as pd
import numpy as np

st.title("Weekly Payroll Calculator with Closers & Enrollers")

# File uploads
hubspot_file = st.file_uploader("Upload HubSpot Deal Tracker CSV")
closer_hours_file = st.file_uploader("Upload Closer Timesheet CSV")
enroller_hours_file = st.file_uploader("Upload Enroller Timesheet CSV")

if hubspot_file and closer_hours_file and enroller_hours_file:
    # Load HubSpot Data
    hubspot_df = pd.read_csv(hubspot_file)
    hubspot_df.columns = hubspot_df.columns.str.strip()
    hubspot_df['CLOSER'] = hubspot_df['CLOSER'].str.strip().str.title()
    hubspot_df['ENROLLER'] = hubspot_df['ENROLLER'].astype(str).str.strip().str.title()
    hubspot_df['DATE'] = pd.to_datetime(hubspot_df['DATE'], errors='coerce')

    # Load Closer Hours
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
    closer_df = closer_df[['Rep', 'Man Hours']].rename(columns={'Rep': 'Agent'})

    # Load Enroller Hours
    enroller_df = pd.read_csv(enroller_hours_file)
    enroller_df.columns = enroller_df.columns.str.strip()
    enroller_df['Rep'] = enroller_df['Rep'].str.strip().str.title()
    enroller_df['Man Hours'] = enroller_df['Man Hours'].astype(str).apply(parse_and_round_up)
    enroller_df = enroller_df[['Rep', 'Man Hours']].rename(columns={'Rep': 'Agent'})

    # Closers Payroll Calculation
    deal_counts = hubspot_df['CLOSER'].value_counts().reset_index()
    deal_counts.columns = ['Agent', 'Deal Count']
    saturday_deals = hubspot_df[hubspot_df['DATE'] == '2025-08-02'].groupby('CLOSER').size().reset_index(name='Saturday Deals')
    saturday_deals['CLOSER'] = saturday_deals['CLOSER'].str.strip().str.title()

    # First Deal of the Day Fix
    hubspot_df['Deal Date'] = hubspot_df['DATE'].dt.date
    first_deals = hubspot_df.sort_values(by='DATE').drop_duplicates(subset=['Deal Date'], keep='first')
    first_deal_bonus = first_deals['CLOSER'].value_counts().reset_index()
    first_deal_bonus.columns = ['Agent', 'First Deal Bonus Count']

    closers = closer_df.merge(deal_counts, on='Agent', how='left').fillna(0)
    closers = closers.merge(saturday_deals.rename(columns={'CLOSER': 'Agent'}), on='Agent', how='left').fillna(0)
    closers = closers.merge(first_deal_bonus, on='Agent', how='left').fillna(0)

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
    closers['Total Deal Pay'] = closers['Regular Deals Pay'] + closers['Saturday Deals Pay']
    closers['Hours Bonus'] = closers['Man Hours'].apply(hours_bonus)
    closers['First Deal Bonus'] = closers['First Deal Bonus Count'] * 25
    closers['Total Bonus Pay'] = closers['Hours Bonus'] + closers['First Deal Bonus']
    closers['Manual Bonus'] = 0
    closers['Total Pay'] = closers['Hourly Pay'] + closers['Total Deal Pay'] + closers['Total Bonus Pay'] + closers['Manual Bonus']
    closers['CPA'] = closers.apply(lambda row: row['Total Pay'] / row['Deal Count'] if row['Deal Count'] > 0 else 0, axis=1)

    # Enrollers Payroll Calculation
    enroller_submissions = hubspot_df['ENROLLER'].value_counts().reset_index()
    enroller_submissions.columns = ['Agent', 'Submitted Deals']

    enrollers = enroller_df.merge(enroller_submissions, on='Agent', how='left').fillna(0)
    enrollers['Hourly Rate'] = 18
    enrollers['Hourly Pay'] = enrollers['Man Hours'] * enrollers['Hourly Rate']
    enrollers['Submitted Deals Pay'] = enrollers['Submitted Deals'] * 5
    enrollers['Manual Bonus'] = 0
    enrollers['Total Pay'] = enrollers['Hourly Pay'] + enrollers['Submitted Deals Pay'] + enrollers['Manual Bonus']
    enrollers['CPA'] = enrollers.apply(lambda row: row['Total Pay'] / row['Submitted Deals'] if row['Submitted Deals'] > 0 else 0, axis=1)

    # Combine Closers and Enrollers into One Export
    closers_export = closers[['Agent', 'Deal Count', 'Man Hours', 'Hourly Rate', 'Hourly Pay',
                              'Total Deal Pay', 'Total Bonus Pay', 'Manual Bonus', 'Total Pay', 'CPA']]
    enrollers_export = enrollers[['Agent', 'Submitted Deals', 'Man Hours', 'Hourly Rate', 'Hourly Pay',
                                  'Submitted Deals Pay', 'Manual Bonus', 'Total Pay', 'CPA']]
    enrollers_export = enrollers_export.rename(columns={'Submitted Deals': 'Deal Count', 'Submitted Deals Pay': 'Total Deal Pay'})

    combined_export = pd.concat([closers_export, enrollers_export], ignore_index=True, sort=False).fillna(0)

    # Display Combined Table
    st.subheader("Payroll Summary (Closers + Enrollers)")
    st.dataframe(combined_export.round(2))

    # Download Combined CSV
    st.download_button("Download Payroll CSV", combined_export.to_csv(index=False).encode('utf-8'), "Payroll_Summary.csv", "text/csv")

else:
    st.warning("Please upload all three CSV files to proceed.")
