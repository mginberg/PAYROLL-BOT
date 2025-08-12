import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from fuzzywuzzy import process
from datetime import date, timedelta

st.title("Weekly Payroll Calculator with Closers & Enrollers")

# ---------- Helpers ----------
def parse_and_round_up(duration_str: str) -> float:
    try:
        h, m, s = map(int, str(duration_str).split(':'))
        total_hours = h + m/60 + s/3600
        return float(np.ceil(total_hours))
    except Exception:
        return 0.0

def fuzzy_match(name: str, name_list, cutoff=80) -> str:
    match = process.extractOne(name, name_list, score_cutoff=cutoff)
    return match[0] if match else name

def determine_hourly_rate(deals: int) -> int:
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

def hours_bonus(hours: float) -> int:
    if hours >= 60:
        return 100
    elif hours >= 50:
        return 75
    elif hours >= 40:
        return 50
    else:
        return 0

def most_recent_saturday(dates: pd.Series) -> date:
    only_dates = pd.to_datetime(dates, errors="coerce").dt.date.dropna()
    saturdays = only_dates[only_dates.map(lambda d: d.weekday() == 5)]
    if len(saturdays) > 0:
        return max(saturdays)
    today = date.today()
    # previous Saturday
    return today - timedelta(days=(today.weekday() - 5) % 7 or 7)

# ---------- Inputs ----------
hubspot_file = st.file_uploader("Upload HubSpot Deal Tracker CSV")
closer_hours_file = st.file_uploader("Upload Closer Timesheet CSV")
enroller_hours_file = st.file_uploader("Upload Enroller Timesheet CSV")

if not (hubspot_file and closer_hours_file and enroller_hours_file):
    st.warning("Please upload all three CSV files to proceed.")
    st.stop()

# ---------- Load Data ----------
hubspot_df = pd.read_csv(hubspot_file)
hubspot_df.columns = hubspot_df.columns.str.strip()
hubspot_df['CLOSER'] = hubspot_df['CLOSER'].astype(str).str.strip().str.title()
hubspot_df['ENROLLER'] = hubspot_df['ENROLLER'].astype(str).str.strip().str.title()
hubspot_df['DATE'] = pd.to_datetime(hubspot_df['DATE'], errors='coerce')
hubspot_df['DealDate'] = hubspot_df['DATE'].dt.date  # date-only for comparisons

# Saturday date selector (defaults to most recent Saturday in the data or last Saturday)
default_sat = most_recent_saturday(hubspot_df['DATE'])
sat_date = st.date_input("Saturday date for $50 per-deal pay", value=default_sat)

# Closer hours
closer_df = pd.read_csv(closer_hours_file)
closer_df.columns = closer_df.columns.str.strip()
closer_df['Rep'] = closer_df['Rep'].astype(str).str.strip().str.title()
closer_df['Man Hours'] = closer_df['Man Hours'].astype(str).apply(parse_and_round_up)

# Enroller hours
enroller_df = pd.read_csv(enroller_hours_file)
enroller_df.columns = enroller_df.columns.str.strip()
enroller_df['Rep'] = enroller_df['Rep'].astype(str).str.strip().str.title()
enroller_df['Man Hours'] = enroller_df['Man Hours'].astype(str).apply(parse_and_round_up)

# ---------- Fuzzy Matching (Closers) ----------
# Map hours -> hubspot closer list
closer_df['Agent'] = closer_df['Rep'].apply(lambda x: fuzzy_match(x, hubspot_df['CLOSER'].unique()))
# Map hubspot closer -> hours list (ensures union alignment)
hubspot_df['Matched Agent'] = hubspot_df['CLOSER'].apply(lambda x: fuzzy_match(x, closer_df['Agent'].unique()))

# Deal counts (stable Agent column)
deal_counts = (
    hubspot_df['Matched Agent']
    .value_counts()
    .rename_axis('Agent')
    .reset_index(name='Deal Count')
)

# Saturday deals (compare on date only)
saturday_deals = (
    hubspot_df.loc[hubspot_df['DealDate'] == sat_date, 'Matched Agent']
    .value_counts()
    .rename_axis('Agent')
    .reset_index(name='Saturday Deals')
)

# First deal of each day (company-wide)
first_per_date = (
    hubspot_df.sort_values(by='DATE')
              .drop_duplicates(subset=['DealDate'], keep='first')
)
first_deal_bonus = (
    first_per_date['Matched Agent']
    .value_counts()
    .rename_axis('Agent')
    .reset_index(name='First Deal Bonus Count')
)

# Build base list = union of hours & deals; ensures agents with hours but 0 deals appear
base_closers = pd.DataFrame(
    sorted(set(closer_df['Agent']).union(set(deal_counts['Agent']))),
    columns=['Agent']
)

# Merge in hours and counts
closers = base_closers.merge(closer_df[['Agent', 'Man Hours']], on='Agent', how='left')
closers = closers.merge(deal_counts, on='Agent', how='left')
closers = closers.merge(saturday_deals, on='Agent', how='left')
closers = closers.merge(first_deal_bonus, on='Agent', how='left')
closers = closers.fillna({'Man Hours': 0, 'Deal Count': 0, 'Saturday Deals': 0, 'First Deal Bonus Count': 0})

# Compute pay pieces
closers['Hourly Rate'] = closers['Deal Count'].apply(determine_hourly_rate)
closers['Hourly Pay'] = closers['Hourly Rate'] * closers['Man Hours']
closers['Regular Deals'] = closers['Deal Count'] - closers['Saturday Deals']
closers['Regular Deals Pay'] = closers['Regular Deals'] * 35
closers['Saturday Deals Pay'] = closers['Saturday Deals'] * 50
closers['Hours Bonus'] = closers['Man Hours'].apply(hours_bonus)
closers['First Deal Bonus'] = closers['First Deal Bonus Count'] * 25

# ---------- Fuzzy Matching (Enrollers) ----------
# Align enroller hours names to hubspot enroller names
enroller_df['Agent'] = enroller_df['Rep'].apply(lambda x: fuzzy_match(x, hubspot_df['ENROLLER'].unique()))
# Count submissions from hubspot
enroller_submissions = (
    hubspot_df['ENROLLER']
    .value_counts()
    .rename_axis('Agent')
    .reset_index(name='Submitted Deals')
)

# Base enrollers = union of hours & submissions
base_enrollers = pd.DataFrame(
    sorted(set(enroller_df['Agent']).union(set(enroller_submissions['Agent']))),
    columns=['Agent']
)
enrollers = base_enrollers.merge(enroller_df[['Agent', 'Man Hours']], on='Agent', how='left')
enrollers = enrollers.merge(enroller_submissions, on='Agent', how='left')
enrollers = enrollers.fillna({'Man Hours': 0, 'Submitted Deals': 0})

# Enroller pay rules
enrollers['Hourly Rate'] = 18
enrollers['Hourly Pay'] = enrollers['Man Hours'] * enrollers['Hourly Rate']
enrollers['Regular Deals Pay'] = enrollers['Submitted Deals'] * 5
enrollers['Saturday Deals Pay'] = 0
enrollers['Hours Bonus'] = 0
enrollers['First Deal Bonus'] = 0
enrollers['Deal Count'] = enrollers['Submitted Deals']  # for unified CPA column

# ---------- Export Frames ----------
closers_export = closers[['Agent', 'Deal Count', 'Man Hours', 'Hourly Rate', 'Hourly Pay',
                          'Regular Deals Pay', 'Saturday Deals Pay', 'Hours Bonus',
                          'First Deal Bonus']].copy()
closers_export['Manual Bonus'] = 0
closers_export['$25 Bonus Count'] = 0
closers_export['$50 Bonus Count'] = 0

enrollers_export = enrollers[['Agent', 'Deal Count', 'Man Hours', 'Hourly Rate', 'Hourly Pay',
                              'Regular Deals Pay', 'Saturday Deals Pay', 'Hours Bonus',
                              'First Deal Bonus']].copy()
enrollers_export['Manual Bonus'] = 0
enrollers_export['$25 Bonus Count'] = 0
enrollers_export['$50 Bonus Count'] = 0

# Sort each section alphabetically by Agent
closers_export = closers_export.sort_values(by='Agent').reset_index(drop=True)
enrollers_export = enrollers_export.sort_values(by='Agent').reset_index(drop=True)

# Combine sections (Closers first, then Enrollers)
combined_export = pd.concat([closers_export, enrollers_export], ignore_index=True).fillna(0)

# ---------- XLSX with Formulas ----------
output = BytesIO()
wb = Workbook()
ws = wb.active
ws.title = "Payroll Summary"

headers = [
    'Agent', 'Deal Count', 'Man Hours', 'Hourly Rate', 'Hourly Pay',
    'Regular Deals Pay', 'Saturday Deals Pay', 'Hours Bonus', 'First Deal Bonus',
    'Manual Bonus', '$25 Bonus Count', '$50 Bonus Count', 'Total Pay', 'CPA'
]
ws.append(headers)

for idx, row in combined_export.iterrows():
    row_num = idx + 2  # header is row 1
    ws.append([
        row['Agent'],
        row['Deal Count'],
        row['Man Hours'],
        row['Hourly Rate'],
        row['Hourly Pay'],
        row['Regular Deals Pay'],
        row['Saturday Deals Pay'],
        row['Hours Bonus'],
        row['First Deal Bonus'],
        '',  # Manual Bonus (editable)
        '',  # $25 Bonus Count (editable)
        '',  # $50 Bonus Count (editable)
        f"=E{row_num}+F{row_num}+G{row_num}+H{row_num}+I{row_num}+J{row_num}+K{row_num}*25+L{row_num}*50",  # Total Pay
        f"=M{row_num}/B{row_num}" if row['Deal Count'] > 0 else "0"  # CPA
    ])

# Overall CPA at bottom
total_rows = len(combined_export) + 2
ws[f"M{total_rows}"] = "Overall CPA:"
ws[f"N{total_rows}"] = f"=SUM(M2:M{total_rows-1})/SUM(B2:B{total_rows-1})"

wb.save(output)

st.download_button(
    "Download Payroll XLSX",
    output.getvalue(),
    "Payroll_Summary.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# On-screen preview (optional)
st.subheader("Preview (rounded)")
st.dataframe(combined_export.round(2))
st.caption(f"Saturday date used for $50 per-deal pay: {sat_date.isoformat()}")
