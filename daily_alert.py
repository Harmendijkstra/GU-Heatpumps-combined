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
from sending_messages import send_email, send_teams_message

sys.path.append('knmi/')
from knmi import get_hour_data_dataframe
station_deelen = 275

# Specify global variables
bRunTwoDaysBefore = True # This variable needs to be True if the script is runned automatically, to get the data two days before. Note that we can not check the data from yesterday as the knmi data will not be available yet.
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

if bRunTwoDaysBefore:
    time_now = datetime.now()
    two_days_previous_date  = time_now - timedelta(days=2) # Go back 2 days
    two_days_previous = two_days_previous_date.strftime('%Y-%m-%d')
    previous_day_date = time_now - timedelta(days=1) # Go back 1 day
    previous_day = previous_day_date.strftime('%Y-%m-%d')
else:
    # Below is the option to set the day manually
    two_days_previous = '2024-10-14'
    previous_day = '2024-10-15'



def check_nans(df, ignore_columns):
    # Filter out the columns to ignore
    columns_to_check = [col for col in df.columns if col[1] not in ignore_columns]
    
    # Check for NaN values in the remaining columns
    nan_columns = [col for col in columns_to_check if df[col].isna().any()]
    
    # Generate the message
    if not nan_columns:
        return "None"
    else:
        nan_columns_str = '\n\n'.join([f"{col[0]}: ({', '.join(map(str, col[1:]))})" for col in nan_columns])
        nan_columns_email_str = '\n'.join([f"{col[0]}: ({', '.join(map(str, col[1:]))})" for col in nan_columns])
        message = f"NaN values at:\n\n{nan_columns_str}"
        email_message = f"NaN values at:\n{nan_columns_email_str}"
        return message, email_message


# Read data into pickles
process_and_save(pBase, pInput, pPickles, sMeetsetFolder)

dfRaw = load_data(pPickles, sMeetset=sMeetsetFolder, sDateStart = two_days_previous, sDateEnd = previous_day)
df, dfHeaders = flatten_data(dfRaw, bStatus=False, ignore_multi_index_differences=True)
df = combine_raw_columns(df)
df = combine_and_sync_rows(df)
df = add_hours_based_on_dst(df, two_days_previous, previous_day)
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

# Fill missing weather data with KNMI data:
# Construct start and end time for knmi data
start_time = (df_1hr.index[0] - timedelta(days=1)).strftime('%Y%m%d%H')
end_time = (df_1hr.index[-1] + timedelta(days=1)).strftime('%Y%m%d%H')
df = get_hour_data_dataframe(stations=[station_deelen], start=start_time, end=end_time , variables=['T','P', 'U'])
# Fill NaN values with KNMI data
temp_air_filled = (df['T'] / 10).reindex(df_1hr.index)
air_pres_filled = (df['P'] / 10).reindex(df_1hr.index)
rel_hum_filled = (df['U']).reindex(df_1hr.index)
df_1hr['Weather Temp Air'] = df_1hr['Weather Temp Air'].combine_first(temp_air_filled)
df_1hr['Weather Abs Air Pressure'] = df_1hr['Weather Abs Air Pressure'].combine_first(air_pres_filled)
df_1hr['Weather Rel Humidity'] = df_1hr['Weather Rel Humidity'].combine_first(rel_hum_filled)


dictHeaderMapping = hHP.genHeaders(df_1min.columns)
df_1hr_newheaders = create_output_dataframe(df_1hr, dictHeaderMapping, is_hourly=True)

# Columns to ignore
ignore_columns = ['V_gas_br', 'Pe_WP1', 'Pe_WP2', 'COP_fabr1', 'COP_fabr2', 'COP_fabr'] # Columns to ignore when checking for NaN values
message, email_message = check_nans(df_1hr_newheaders, ignore_columns)
# Check max_consecutive_count
max_consecutive_issues = {k: v for k, v in max_consecutive_count.items() if v > 1}
# Construct the email content
email_subject = "Heatpump Monitoring"
body = ""
email_body = ""

if max_consecutive_issues:
    max_consecutive_str = '\n'.join([f"{k}: {v}" for k, v in max_consecutive_issues.items()])
    body += f"Columns with consecutive outliers:\n{max_consecutive_str}\n\n"
    email_body += f"Columns with consecutive outliers:\n{max_consecutive_str}\n\n"

if message != "None":
    body += f"{message}\n\n"
    email_body = f"{email_message}"

# Send the email if there are any issues
if body:
    send_email(email_subject, email_body, "robert.mellema@dnv.com", sMeetsetFolder, two_days_previous) # With DNV, one needs to 'allow sender' within review of quarantined emails
    send_teams_message(body, sMeetsetFolder, two_days_previous)
