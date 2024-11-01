import os
from os.path import join as cmb
import pandas as pd
import excelMethods as xlm
import time


def create_output_dataframe(df, lstHeaderMapping, is_hourly):
    # Check if all columns in the header mapping exist in the DataFrame
    print("\nCreate output dataframe...")
    missing_columns = [col for col in lstHeaderMapping if col not in df.columns]
    if missing_columns:
        raise ValueError(f"The following columns are missing in the DataFrame: {missing_columns}")

    # Select the columns and create a new DataFrame
    selected_columns = {}
    for col in lstHeaderMapping:
        if col in df.columns:
            if isinstance(lstHeaderMapping[col], dict):
                selected_columns[col] = lstHeaderMapping[col]["hourly_data"] if is_hourly else lstHeaderMapping[col]["minute_data"]
            else:
                selected_columns[col] = lstHeaderMapping[col]

    new_df = df[list(selected_columns.keys())].copy()

    # Extract 'Datum' and 'Tijd' from 'Adjusted Timestamp'
    new_df['Datum'] = df['Adjusted Timestamp'].dt.strftime('%d-%m-%Y')
    new_df['Tijd'] = df['Adjusted Timestamp'].dt.strftime('%H:%M:%S')

    # Create a MultiIndex for the columns
    multi_index_tuples = [('Datum', '', ''), ('Tijd', '', '')] + list(selected_columns.values())
    multi_index = pd.MultiIndex.from_tuples(multi_index_tuples, names=['MV', 'Description', 'Unit'])

    # Reorder columns to place 'Datum' and 'Tijd' at the beginning
    new_df = new_df[['Datum', 'Tijd'] + list(selected_columns.keys())]

    new_df.columns = multi_index

    return new_df

def get_decimal_places_mapping(input_data, input_type='dict', is_hourly=False):
    decimal_places = {}

    def process_item(mv_key, dimension, key):
        if dimension in ['m3/h', 'kWh', 'bara', 'm3(n)/h', 'kJ/kg', 'kg/h', 'kJ/s', 'kJ/s']:
            decimal_places[mv_key] = 3
        elif dimension in ['\u00b0C', 'hPa']:
            decimal_places[mv_key] = 2
        elif dimension in ['%', 'kW', 'l/h', 'W']:
            decimal_places[mv_key] = 0
        elif dimension == '-':
            # Skip adding to decimal_places mapping
            return
        else:
            if key not in ['Datum', 'Tijd']:
                print(f"The key: {key} has not a specified amount of decimals, please specify it!")
            decimal_places[mv_key] = 2  # Default to 2 decimals if not specified

    if input_type == 'dict':
        for key, value in input_data.items():
            if type(value)==dict:
                if is_hourly:
                    mv_key = value["hourly_data"][0]
                    dimension = value["hourly_data"][2]
                else:
                    mv_key = value["minute_data"][0]
                    dimension = value["minute_data"][2]
            else:
                mv_key = value[0]
                dimension = value[2]
            process_item(mv_key, dimension, key)

    elif input_type == 'df':
        for col in input_data.columns:
            mv_key = col[0]
            dimension = col[2]
            process_item(mv_key, dimension, mv_key)

    else:
        raise ValueError("Invalid input_type. Expected 'dict' or 'df'.")

    return decimal_places

def save_dataframe_with_dates(df, lstHeaderMapping, folder_dir, prefix='', header_input_type='dict'):
    print("Save dataframe to xlsx...")
    df_output = df.copy()
    
    # Ensure the folder exists
    if not os.path.exists(folder_dir):
        os.makedirs(folder_dir)

    # Extract start and end dates from the 'Datum' column
    start_date = df_output[('Datum', '', '')].iloc[0].replace('/', '-')
    end_date = df_output[('Datum', '', '')].iloc[-1].replace('/', '-')

    # Determine the file name based on the dates
    if start_date == end_date:
        file_name = f"{prefix}Energy Balance - {start_date}.xlsx"
    else:
        file_name = f"{prefix}Energy Balance - {start_date} - {end_date}.xlsx"

    # Full path to save the file
    file_path = os.path.join(folder_dir, file_name)

    # Get the decimal places mapping
    if 'min' in prefix:
        is_hourly = False
    elif 'hour' in prefix:
        is_hourly = True
    else:   # raise an error
        raise ValueError("The prefix must contain 'min' or 'hour'.")
    if header_input_type == 'df':
        decimal_places_mapping = get_decimal_places_mapping(df, input_type='df', is_hourly=is_hourly)
    elif header_input_type == 'dict':
        decimal_places_mapping = get_decimal_places_mapping(lstHeaderMapping, input_type='dict', is_hourly=is_hourly)

    # Round the DataFrame columns based on the decimal places mapping
    for col in df_output.columns:
        col_name = col[0]  # Assuming the first level of the MultiIndex is the key
        if col_name in decimal_places_mapping:
            df_output[col] = df_output[col].round(decimal_places_mapping[col_name])

    # Save the DataFrame to an Excel file using pd.ExcelWriter
    # If file exist, delete it
    if os.path.exists(file_path):
        os.remove(file_path)
    # Save the DataFrame to an Excel file using pd.ExcelWriter
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df_output.to_excel(writer, index=True, index_label=None, sheet_name='Sheet1')
    print(f"DataFrame saved to {file_path}")
    return file_path

#Note below is a function special made for this project
def convert_excel_output(folder_dir, change_files):
    max_retries = 5
    retry_delay = 1  # seconds

    xlApp = xlm.xlOpen(bXLvisible=1)
    sDecimalSeparator = xlApp.International[2]

    for sFile in change_files:
        retry_count = 0
        while retry_count < max_retries:
            try:
                xlWb = xlm.openWorkbook(xlApp, cmb(folder_dir, sFile))
                break
            except Exception as e:
                retry_count += 1
                if retry_count < max_retries:
                    print(f"Error opening workbook: {e}. Retrying in {retry_delay} seconds...")
                    time.sleep(retry_delay)
                else:
                    print(f"Failed to open workbook after {max_retries} retries: {e}")
                    raise

        if xlWb:
            try:
                xlSheet = xlWb.Sheets("Sheet1")
                xlm.autofitColumns(xlSheet)
                xlm.removeRow(xlSheet, 4)
                xlSheet.Range("D4").Select()
                xlApp.ActiveWindow.FreezePanes = True
                xlSheet.Range("D:E").NumberFormat = f'0{sDecimalSeparator}0%'
                xlm.closeWorkbook(xlApp, xlWb, bSave=True)
            except Exception as e:
                print(f"Error processing workbook: {e}")
                xlm.closeWorkbook(xlApp, xlWb, bSave=False)

    xlm.xlClose(xlApp)