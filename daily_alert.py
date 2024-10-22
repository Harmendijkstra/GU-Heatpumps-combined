import os
import sys
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
from os.path import join as cmb
import headerMappingsHP as hHP
from HPimport import (
    process_and_save,
    load_data,
    flatten_data,
    combine_raw_columns,
    combine_and_sync_rows,
    add_hours_based_on_dst,
    check_monotonic_and_fill_gaps,
    make_totalizers_monotonic,
    convert_to_1_minute_data,
    interpolate_nans,
    sortColumns,
    remove_outliers,
    calc_weithed_mean_flow,
    add_cop_values,
    convert_to_1_hour_data,
    create_output_dataframe
)
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import requests
import json

# Specify global variables
bRunPreviousDay = True # This variable needs to be True if the script is runned automatically, to get the previous day data 
maximum_interpolate_nans = 20 # Number of NaNs to interpolate in a row before setting the rest to nan


# For all variables there is a min and max values setted in minmax_cols.xlsx, but TgasIn and PgasIn are also set here:
TgasIn_min = -5.0
TgasIn_max = 100.0
PgasIn_min = -2.0
PgasIn_max = 70.0
atmospheric_pressure = 1.023  # Atmospheric pressure in bar
minimum_atmospheric_pressure = 700 # Minimum pressure in mbar, below which the pressure is considered to be invalid and atmospheric_pressure is used


pBase = os.getcwd()
pParentDir = os.path.dirname(pBase)
pInput = cmb(pParentDir,'Collected Data')
pPickles = cmb(pBase,'ImportedPickles')

# By default, the script will use the first two arguments as the meetset folder and location
# One can use this by running the script from the command line with the meetset folder and location as arguments. For example python daily_alert.py "Meetset1-Deventer" "Deventer"
if len(sys.argv) > 2:
    sMeetsetFolder = sys.argv[1]
    location = sys.argv[2]
else:
    # Else use the default values, that can be set here manually:
    sMeetsetFolder = 'Meetset1-Deventer'
    location = 'Deventer' # Used for Automatic excel calculations within Word documents

if bRunPreviousDay:
    time_now = datetime.now()
    previousday_date  = time_now - timedelta(days=1) # Go back 1 day to the previous day
    previous_day = previousday_date.strftime('%Y-%m-%d')
    this_day = time_now.strftime('%Y-%m-%d')
else:
    # Below is the option to set the day manually
    previous_day = '2024-10-14'
    this_day = '2024-10-15'



def check_nans(df, ignore_columns):
    # Filter out the columns to ignore
    columns_to_check = [col for col in df.columns if col[1] not in ignore_columns]
    
    # Check for NaN values in the remaining columns
    nan_columns = [col for col in columns_to_check if df[col].isna().any()]
    
    # Generate the message
    if not nan_columns:
        return "None"
    else:
        nan_columns_str = ', '.join([f"{col[0]} ({col[1]}, {col[2]})" for col in nan_columns])
        return f"The following columns contain NaN values: {nan_columns_str}."

# Function to send email
def send_email_with_html(subject, body, email_receiver, sMeetsetFolder, this_day):
    email_sender ="monitoringheatpump@gmail.com"
    email_password = "crhr pwiw pwky icsf"

    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Heatpump Monitoring"
    msg['From'] = email_sender
    msg['To'] = email_receiver

    # html = '<html><body><p>Hi, I have the following alerts for you!</p></body></html>'
    intro_message = f"Hello colleagues,\n\nHeatpump monitoring saw a problem in the daily data for {sMeetsetFolder} at {this_day}.\n"
    end_message = "\n\nGreetings,\nHeatpump Monitoring System"
    full_body = intro_message + body + end_message

    # Create the HTML part of the email
    html = f'<html><body><p>{full_body.replace("\n", "<br>")}</p></body></html>'
    part2 = MIMEText(html, 'html')

    msg.attach(part2)

    # Send the message via gmail's regular server, over SSL - passwords are being sent, afterall
    s = smtplib.SMTP_SSL('smtp.gmail.com')
    # uncomment if interested in the actual smtp conversation
    # s.set_debuglevel(1)
    # do the smtp auth; sends ehlo if it hasn't been sent already
    s.login(email_sender, email_password)

    s.sendmail(email_sender, email_receiver, msg.as_string())
    s.quit()


def send_email(subject, body, email_receiver, sMeetsetFolder, this_day):
    email_sender = "monitoringheatpump@gmail.com"
    email_password = "crhr pwiw pwky icsf"  # Replace with your app password if 2FA is enabled

    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Heatpump Monitoring"
    msg['From'] = email_sender
    msg['To'] = email_receiver

    intro_message = f"Hello colleagues,\n\nHeatpump monitoring saw a problem in the daily data for {sMeetsetFolder} at {this_day}.\n"
    end_message = "\n\nGreetings,\nHeatpump Monitoring System"
    full_body = intro_message + body + end_message

    # Attach the plain text part to the email
    part1 = MIMEText(full_body, 'plain')
    msg.attach(part1)

    try:
        # Send the message via Gmail's regular server, using STARTTLS
        s = smtplib.SMTP('smtp.gmail.com', 587)
        # s.set_debuglevel(1)  # Enable debug output
        s.ehlo()  # Identify ourselves to the SMTP server
        s.starttls()  # Secure the connection
        s.ehlo()  # Re-identify ourselves as an encrypted connection
        s.login(email_sender, email_password)
        s.sendmail(email_sender, email_receiver, msg.as_string())
        s.quit()
        print("Email sent successfully")
    except Exception as e:
        print(f"Failed to send email: {e}")

def send_teams_message(body, sMeetsetFolder, this_day):
    webhook_url = "https://dnv.webhook.office.com/webhookb2/de5c61a7-826f-4f87-9c9f-93f5366aa625@adf10e2b-b6e9-41d6-be2f-c12bb566019c/IncomingWebhook/20e3dad24a824de3ac2695892b1c8fab/5fcef47d-1ed1-4d15-92b3-dc1169d4a35e/V2Yeeb2ayiGQxihAZM7LRmzLJLWroIrbt4M-0Ntz_kaEg1"
    intro_message = f"Hello colleagues,\n\nHeatpump monitoring saw a problem in the daily data for {sMeetsetFolder} at {this_day}.\n"
    end_message = "\n\nGreetings,\nHeatpump Monitoring System"
    full_message = intro_message + body + end_message

    headers = {'Content-Type': 'application/json'}
    payload = {
        "text": full_message
    }
    
    response = requests.post(webhook_url, headers=headers, data=json.dumps(payload))
    
    if response.status_code == 200:
        print("Message sent to Teams successfully.")
    else:
        print(f"Failed to send message to Teams: {response.status_code}, {response.text}")


def send_sms_via_email(phone_number, carrier_domain, subject, message, email_sender, email_password):
    # Construct the email
    sms_recipient = f"{phone_number}@{carrier_domain}"
    msg = MIMEText(message)
    msg['From'] = email_sender
    msg['To'] = sms_recipient
    msg['Subject'] = subject

    try:
        # Send the message via Gmail's SMTP server
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(email_sender, email_password)
        s.sendmail(email_sender, sms_recipient, msg.as_string())
        s.quit()
        print("SMS sent successfully via email")
    except Exception as e:
        print(f"Failed to send SMS: {e}")


# Read data into pickles
process_and_save(pBase, pInput, pPickles, sMeetsetFolder)

dfRaw = load_data(pPickles, sMeetset=sMeetsetFolder, sDateStart = previous_day, sDateEnd = this_day)
df, dfHeaders = flatten_data(dfRaw, bStatus=False, ignore_multi_index_differences=True)
df = combine_raw_columns(df)
df = combine_and_sync_rows(df)
df = add_hours_based_on_dst(df, previous_day, this_day)
df.set_index('Adjusted Timestamp', drop=False, inplace=True)
df = check_monotonic_and_fill_gaps(df, freq='15s')
df = interpolate_nans(df, nLimit=maximum_interpolate_nans)

# First interpolate, then make sum of Eastron power. Otherwise, some interpolated values are missing in the sum.
if 'Weather Abs Air Pressure' in df.columns:
    # Convert pgasin from barg to bara, but if air pressure is below minimum_atmospheric_pressure mbar, add atmospheric_pressure to the value instead of reading the air pressure
    df['PgasIn'] = np.where(
        df['Weather Abs Air Pressure'].fillna(atmospheric_pressure) < minimum_atmospheric_pressure,
        atmospheric_pressure + df['PgasIn'],
        df['Weather Abs Air Pressure'].fillna(atmospheric_pressure).div(1000) + df['PgasIn']
    )
if 'Eastron01 Total Power' in df.columns and 'Eastron02 Total Power' in df.columns:
    df['Eastron Total Power'] = df[['Eastron01 Total Power', 'Eastron02 Total Power']].sum(axis=1)
    rowsMaskNAN = df[['Eastron01 Total Power', 'Eastron02 Total Power']].isna().all(axis=1)
    df.loc[df[rowsMaskNAN].index, 'Eastron Total Power'] = np.nan


# Check on outlier values:
# TgasIn: Outliers for temperature and PgasIn: outliers for incoming pressure
rowsMask = df[~df['TgasIn'].between(TgasIn_min, TgasIn_max)].index
df.loc[rowsMask, 'TgasIn'] = np.nan
rowsMask = df[~df['PgasIn'].between(PgasIn_min, PgasIn_max)].index
df.loc[rowsMask, 'PgasIn'] = np.nan

# Handle totalizers, make them consistent and then use for 1min output.
dict_totalizer = {
    'Itron Gas volume 1': ['Itron Gas volume 1'],
    'Belimo01 FlowTotalL': ['Belimo01 FlowTotalL'],
    'Belimo02 FlowTotalL': ['Belimo02 FlowTotalL'],
    'Belimo03 FlowTotalL': ['Belimo03 FlowTotalL'],
    'BelimoValve FlowTotalL': ['BelimoValve FlowTotalL'],
    # Add 'Hager Total Energy' as well? Currently not used, and no errors..
    }
dict_totalizer = {k: dict_totalizer[k] for k in dict_totalizer if k in df.columns}
colsTotalizer = [item for sublist in list(dict_totalizer.values()) for item in sublist]
df = make_totalizers_monotonic(df, colsTotalizer)

# Remove lines where no data was logged, so check a few basic columns are all nan:
dropIndices = df[df[['PgasIn','TgasIn']].isna().all(axis=1)].index        

df.drop(dropIndices.unique(), axis=0, inplace=True)

# Now work on the 1min output for the energy balance calculation    
df_1min = convert_to_1_minute_data(df, dict_totalizer)
df_1min.set_index('Adjusted Timestamp', drop=False, inplace=True)

df_1min = sortColumns(df_1min, ['Adjusted Timestamp'])
minmax_filepath = cmb(pBase, "minmax_cols.xlsx")
df_minmax = pd.read_excel(minmax_filepath)
df_1min, outliers_count, max_consecutive_count = remove_outliers(df_1min, df_minmax, lstHeaderMapping=hHP.makeAllHeaderMappings())

# Add some columns with weighted mean flow rates
df_1min = calc_weithed_mean_flow(df_1min, col_wm_flow='Q_ket1_wm', col_flow='Belimo01 FlowRate', col_tempout='Belimo01 Temp2 internal', col_tempin='Belimo01 Temp1 external')
df_1min = calc_weithed_mean_flow(df_1min, col_wm_flow='Q_OV_wm', col_flow='Belimo02 FlowRate', col_tempout='Belimo02 Temp2 internal', col_tempin='Belimo02 Temp1 external')
df_1min = calc_weithed_mean_flow(df_1min, col_wm_flow='Q_WP_wm', col_flow='Belimo03 FlowRate', col_tempout='Belimo03 Temp2 internal', col_tempin='Belimo03 Temp1 external')
df_1min = calc_weithed_mean_flow(df_1min, col_wm_flow='Q_klep_wm', col_flow='BelimoValve FlowRate', col_tempout='BelimoValve Temp2 internal', col_tempin='BelimoValve Temp1 external')

df_1min = add_cop_values(df_1min)
df_1hr = convert_to_1_hour_data(df_1min)
dictHeaderMapping = hHP.genHeaders(df_1min.columns)
df_1hr_newheaders = create_output_dataframe(df_1hr, dictHeaderMapping, is_hourly=True)

# Columns to ignore
ignore_columns = ['V_gas_br', 'Pe_WP1', 'Pe_WP2', 'cop_wm1', 'cop_wm2'] # Columns to ignore when checking for NaN values
message = check_nans(df_1hr_newheaders, ignore_columns)
# Check max_consecutive_count
max_consecutive_issues = {k: v for k, v in max_consecutive_count.items() if v > 1}
# Construct the email content
email_subject = "Heatpump Monitoring"
body = ""

if max_consecutive_issues:
    max_consecutive_str = '\n'.join([f"{k}: {v}" for k, v in max_consecutive_issues.items()])
    body += f"Columns with consecutive outliers:\n{max_consecutive_str}\n\n"

if message != "None":
    body += f"{message}\n"

# Send the email if there are any issues
if body:
    send_email(email_subject, body, "harmen.dijkstra@dnv.com", sMeetsetFolder, this_day) # This does not work as DNV is blocking SMTP protocol
    send_teams_message(body, sMeetsetFolder, this_day)
