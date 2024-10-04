# -*- coding: utf-8 -*-
"""
Created on Wed Jun 19 11:14:24 2024

@author: LENLUI
"""

import os
from os.path import join as cmb
import numpy as np
import pandas as pd
from io import StringIO
from datetime import datetime, timedelta
import time
import pytz
import struct
import headerMappingsHP as hHP


def process_and_save(pBase, sMeetset):
    pMeetset = cmb(pInput, sMeetset)
    output_subdir = cmb(pPickles, sMeetset)
    os.makedirs(output_subdir, exist_ok=True)
    print(f'Processing and saving pickles for set {sMeetset}...')

    # Compare ingested file dates in the form of '240607' with lstDone, skip already ingested files.
    lstDone = sorted([f.split('_')[1][2:10].replace('-','') for f in os.listdir(output_subdir) 
                        if (f.endswith('.bz2') and not '_partial' in f)])
    all_files = sorted([f for f in os.listdir(pMeetset) if (f.endswith('.txt') and 
                    not f.split('_')[-2] in lstDone and not '-test' in f and not '-config' in f)])
    daily_data = []
    current_day = None
    lstTimes = []

    for file in all_files:
        print(f' - Reading file {file}..')
        sData, header_df = read_file(cmb(pMeetset, file))
        dfEmpty = create_multi_index_df(header_df)
        tStart = time.time()
        data = process_datafile(sData, dfEmpty, header_df)
        lstTimes.append(time.time()-tStart)

        if not data.empty:
            data[data.columns[0]] = pd.to_datetime(data[data.columns[0]], format='%y/%m/%d %H:%M:%S')
            file_date = data['Timestamp'].iloc[0].iloc[0].date()  # Assuming 'timestamp' column exists and is datetime type
#EDIT python 3.12: iloc[0][0] no longer valid
            if current_day is None:
                current_day = file_date
            if file_date != current_day:
                # Save the accumulated data for the previous day
                daily_df = pd.concat(daily_data)
                daily_df.reset_index(inplace=True, drop=True)
                daily_date = daily_df['Timestamp'].iloc[0].iloc[0].date()
#EDIT python 3.12: iloc[0][0] no longer valid                
                print(f'Saving daily data for date {daily_date}..')
                daily_df = group_and_combine_columns(daily_df)
                daily_df.to_pickle(cmb(output_subdir, f'{sMeetset}_{daily_date}.bz2'))
                if os.path.exists(cmb(output_subdir, f'{sMeetset}_{daily_date}_partial.bz2')):
                    os.remove(cmb(output_subdir, f'{sMeetset}_{daily_date}_partial.bz2'))
                daily_data = []
                current_day = file_date
                print(f'Average processing_datafile execution time was {round(np.mean(lstTimes),2)} sec')
                lstTimes = []
            daily_data.append(data)
    # Handle the last batch of files that do not make up a whole day yet
    if daily_data:
        print(f'Saving partial data for date {file_date}..')
        daily_df = pd.concat(daily_data)
        daily_df.reset_index(inplace=True, drop=True)
        daily_df = group_and_combine_columns(daily_df)
        daily_df.to_pickle(cmb(output_subdir, f'{sMeetset}_{file_date}_partial.bz2'))

def combine_columns(df, cols, combined_col_name):
    """
    Combine a list of columns in a DataFrame, giving precedence to non-NaN values.

    This function takes a DataFrame and a list of column names, and combines these columns into a single column.
    The resulting column will have non-NaN values from the first column in the list, and if a value is NaN, it will
    be replaced by the corresponding non-NaN value from the next column in the list, and so on. The combined column
    is returned as a DataFrame with the specified column name.

    Parameters:
    df (pd.DataFrame): The DataFrame containing the columns to combine.
    cols (list): A list of column names to combine.
    combined_col_name (tuple): The name for the combined column.
    
    Returns:
    pd.Series: A Series with the combined values.
    """
    combined = df[cols[0]]
    for col in cols[1:]:
        combined = combined.combine_first(df[col])
    return pd.DataFrame({combined_col_name: combined})

def group_and_combine_columns(df, combine_column='Signal No.'):
    """
    Group columns by common indices except for the 'Signal No.' level and combine them.
    
    Parameters:
    df (pd.DataFrame): The DataFrame containing the columns to group and combine.
    combine_column (str): The multi index column that should be combined, hence ignored if it has different 
    
    Returns:
    pd.DataFrame: The DataFrame with combined columns.
    """
    # Find the index position of combine_column
    signal_no_index = df.columns.names.index(combine_column)

    # Group columns by common indices except for the signal number
    grouped_columns = {}
    for col in df.columns:
        key = tuple([col[i] for i in range(len(col)) if i != signal_no_index])  # Exclude the signal number index
        if key not in grouped_columns:
            grouped_columns[key] = []
        grouped_columns[key].append(col)

    # Create a list to store the combined columns
    combined_columns = []

    # Combine columns in each group
    for key, cols in grouped_columns.items():
        if len(cols) > 1:
            combined_col_name = key[:signal_no_index] + ('Combined',) + key[signal_no_index:]
            combined_col_df = combine_columns(df, cols, combined_col_name)
        else:
            # Adjust the column name to include 'Combined' in the correct position
            single_col_name = key[:signal_no_index] + ('Combined',) + key[signal_no_index:]
            combined_col_df = df[cols].copy()
            combined_col_df.columns = pd.MultiIndex.from_tuples([single_col_name])
        
        combined_columns.append(combined_col_df)

    # Concatenate all combined columns into a new DataFrame
    combined_df = pd.concat(combined_columns, axis=1)

    # Set the names of the MultiIndex levels
    combined_df.columns.names = df.columns.names

    return combined_df

def read_file(fPath):
    """
    Reads the data rows line by line and processes them into a DataFrame.
    Separately reads the first rows as headers.
    """
    # Open the file in read mode
    with open(fPath, 'r') as file:
        # Read the entire contents of the file into a string
        sData = file.read()
    # Get rid of the trailing separator on each row
    sData = sData.replace(";\n","\n")
    
    # Read the header rows
    header_data = StringIO('\n'.join(sData.split('\n')[:4]))
    header_df = pd.read_csv(header_data, sep=';', header=None)
    
    return sData, header_df


def create_multi_index_df(header_df):
    """
    Creates the multi_index Dataframe from the header rows, returns an empty df with multi-index
    """
    # Create the multi-index
    nSignals = len(header_df.iloc[0])
    # multi_index = pd.MultiIndex.from_arrays([header_df.iloc[1], header_df.iloc[2], header_df.iloc[3], header_df.iloc[0]], names=['Signal Name', 'Unit', 'Tag', 'Signal No.'])
    
    # Below is the option for a MultiIndex that stores [Value, status_bit] for each signal
    multi_indexValues = pd.MultiIndex.from_arrays([header_df.iloc[1], header_df.iloc[2], header_df.iloc[3], header_df.iloc[0], ['Value']*nSignals], names=['Signal Name', 'Unit', 'Tag', 'Signal No.', 'Type'])
    multi_indexStatus = pd.MultiIndex.from_arrays([header_df.iloc[1, 1:], header_df.iloc[2, 1:], header_df.iloc[3, 1:], header_df.iloc[0, 1:], ['Status']*(nSignals-1)], names=['Signal Name', 'Unit', 'Tag', 'Signal No.', 'Type'])
    lst = multi_indexValues.to_flat_index().append(multi_indexStatus.to_flat_index())
    multi_index = pd.MultiIndex.from_tuples(lst, names=['Signal Name', 'Unit', 'Tag', 'Signal No.', 'Type'])
    
    # Initialize an empty DataFrame with the multi-index columns
    dfEmpty = pd.DataFrame(columns=multi_index)
    return dfEmpty

def process_datafile(sData, dfEmpty, header_df):
    """
    Read the sData string, and use the multi-index empty dataframe to create 
    the full dataframe for each file, with a multi-index header
    """    
    # Make sure the timestamp column is set up as multi-index as well
    sTimestampCol = dfEmpty.columns[0]
    rows = []
    data_rows = sData.split('\n')[4:]
    # Determine the lowest signal no. to use in indexing
    # lstSignalNos = [0] + [int(x[3]) for x in dfEmpty.columns[1:]]
    lstSignalNos = header_df.iloc[0]
    # nSignalNoOffset = min() - 1
    
    # Process each data row
    for cnt, row in enumerate(data_rows):
        if row.startswith('D'):
            # print(f'data_rows[{cnt}]')
            parts = row.split(';')
            timestamp = parts[1]
            data_dict = {sTimestampCol: timestamp}
            # Skip row if there are not complete sets of 3 values on each row
            if np.mod(len(parts)-2, 3) != 0:
                print(f'data_row[{cnt}] skipped, because amount of data values is not a multiple of 3 [signal_no, value, status]')
                continue
            for i in range(2, len(parts), 3):
                if i + 2 < len(parts):
                    signal_number = int(parts[i])
                    if parts[i + 1] in ['','T-Mobile  NL','LTE CAT M1','False']:
                        signal_value = np.nan
                    else:
                        signal_value = float(parts[i + 1])
                    status_bit = int(parts[i + 2])
                    # nCol = lstSignalNos.index(signal_number)
                    nCol = lstSignalNos[lstSignalNos == str(signal_number)].index[0]
                    signal_name = header_df.iloc[1, nCol]
                    unit = header_df.iloc[2, nCol]
                    tag = header_df.iloc[3, nCol]
                    # tag = header_df.iloc[3, signal_number - nSignalNoOffset]
                    data_dict[(signal_name, unit, tag, signal_number, 'Value')] = signal_value
                    data_dict[(signal_name, unit, tag, signal_number, 'Status')] = status_bit
            # df = df.append(data_dict, ignore_index=True)
            rows.append(data_dict)
    
    # Create the DataFrame from the list of rows
    df = pd.DataFrame(rows)
    # Set the multi-index columns
    df.columns = pd.MultiIndex.from_tuples(df.columns, names=['Signal Name', 'Unit', 'Tag', 'Signal No.', 'Type'])
    return df


def load_data(sMeetset=None, sDateStart='2024-01-01', sDateEnd='2025-12-31'):
    print(f'\nLoading stored data from {sDateStart} to {sDateEnd}..')
    tStart = datetime.strptime(sDateStart, '%Y-%m-%d').date()
    tEnd = datetime.strptime(sDateEnd, '%Y-%m-%d').date()
    # Only add files for which the YYYY-mm-dd date falls within tStart to tEnd (inclusive)
    lstPickles = [f for f in os.listdir(cmb(pPickles,sMeetset)) if (not f.endswith('partial.bz2') and
                  tStart <= datetime.strptime(f.split('_')[-1][0:10], '%Y-%m-%d').date() <= tEnd) or
                  (f.endswith('partial.bz2') and 
                   tStart <= datetime.strptime(f.split('_')[-2][0:10], '%Y-%m-%d').date() <= tEnd)]
    lstPickleData = []
    for pklfile in lstPickles:
        pkldata = pd.read_pickle(cmb(pPickles,sMeetset,pklfile))
        lstPickleData.append(pkldata)
        print(f'  Loaded file {pklfile}..')
    df = pd.concat(lstPickleData)
    df.reset_index(inplace=True, drop=True)
    return df

def flatten_data(df, bStatus=False, ignore_multi_index_differences=False):
    """
    Group columns by common indices except for the 'Signal No.' level and combine them.
    
    Parameters:
    df (pd.DataFrame): The DataFrame containing the columns to group and combine.
    bStatus (bool): Can be set to True. In that cas wil Status and Value be added as seperate columns
    ignore_multi_index_differences (bool): Can be set to True to ignore differences in the multi index. 
    In that case it would be ignored if there are differences in Unit or Tag.
    
    Returns:
    pd.DataFrame: The DataFrame with combined columns.
    """
    print("Flatten data...")
    if ignore_multi_index_differences:
        df = group_and_combine_columns(df, combine_column='Unit')
        df = group_and_combine_columns(df, combine_column='Tag')
    # Drop the 'Type' level from the MultiIndex columns
    flattened_columns = df.columns.droplevel('Type')
    # Remove duplicates from the flattened columns
    unique_columns = flattened_columns.drop_duplicates()
    # Get all unique signal names
    signal_names = df.columns.get_level_values('Signal Name').unique()
    # Check for each signal name whether it appears only once in the unique columns
    for signal_name in signal_names:
        occurrences = unique_columns[unique_columns.get_level_values('Signal Name') == signal_name]
        count = sum(unique_columns.get_level_values('Signal Name') == signal_name)
        if count > 1:
            raise ValueError(f"There are multiple columns for the signal name '{signal_name}': {count}. Occurrences: {occurrences}")

    if bStatus:
        idxFlatHeader = df.columns.get_level_values(0) + ' (' + df.columns.get_level_values(4) + ')'
    else:
        type_no_index = df.columns.names.index('Type')
        df = df.xs('Value', axis=1, level=type_no_index, drop_level=False)
        idxFlatHeader = df.columns.get_level_values(0)
    dfHeaders = df.columns.to_frame(index=False)
    df.columns = idxFlatHeader
    # df.set_index(idxFlatHeader[0], drop=True, inplace=True)
    return df, dfHeaders

def round_to_nearest_15_seconds(timestamp):
    """Round a timestamp to the nearest 15-second interval."""
    return timestamp.round('15s')

def combine_and_sync_rows(df):
    """Process the DataFrame to round 15sec intervals and fill values."""
    print("Combine and sync rows...")
    # Round timestamps to the nearest 15-second interval
    df['Rounded Timestamp'] = df['Timestamp'].apply(round_to_nearest_15_seconds)
    
    # Group by the rounded timestamps and apply forward fill within each group
    grouped = df.groupby('Rounded Timestamp').apply(lambda group: group.ffill().iloc[-1])
    
    # Reset the index to get a DataFrame
    processed_df = grouped.reset_index(drop=True)
    
    # Drop the original 'Timestamp' column and rename 'Rounded Timestamp' to 'Timestamp'
    processed_df = processed_df.drop(columns=['Timestamp'])
    processed_df = processed_df.rename(columns={'Rounded Timestamp': 'Timestamp'})
    
    return processed_df

def add_hours_based_on_dst(df):
    """
    Add 1 hour or 2 hours to the 'Timestamp' column depending on whether it is Dutch summer time or winter time.
    """
    print("Add Adjusted Timestamp...")
    # Define the Europe/Amsterdam timezone
    amsterdam_tz = pytz.timezone('Europe/Amsterdam')
    
    def adjust_time(timestamp):
        # Localize the timestamp to Europe/Amsterdam timezone
        localized_timestamp = amsterdam_tz.localize(timestamp)
        
        # Check if the timestamp is in DST
        if localized_timestamp.dst() != timedelta(0):
            # Summer time (DST)
            return timestamp + timedelta(hours=2)
        else:
            # Winter time (Standard Time)
            return timestamp + timedelta(hours=1)
    
    # Rename the original 'Timestamp' column
    df = df.rename(columns={'Timestamp': 'Original Timestamp'})
    
    # Apply the adjust_time function to the 'Original Timestamp' column
    df['Adjusted Timestamp'] = df['Original Timestamp'].apply(adjust_time)
    
    # Drop the 'Original Timestamp' column
    df = df.drop(columns=['Original Timestamp'])
    
    return df

def interpolate_columns(df, columns, nLimit=None):
    """Interpolate the specified columns in the DataFrame."""
    for col in columns:
# TODO Check if works everywhere
        # df[col] = df[col].interpolate(method='linear')
        if not 'Timestamp' in col:
            df[col] = df[col].interpolate(method='time', limit=nLimit)
    return df

def process_totalizers(df, totalizer_dict):
    """Process the totalizer columns as specified."""
    for new_col, cols in totalizer_dict.items():
        # Interpolate the columns in the list
        df = interpolate_columns(df, cols)

        # Set NaN values at the beginning and end of each column to the first and last non-NaN values
        for col in cols:
            first_valid_index = df[col].first_valid_index()
            last_valid_index = df[col].last_valid_index()
            
            if first_valid_index is not None:
                # first_valid_loc = df.index.get_loc(first_valid_index)
                # df[col].iloc[:first_valid_loc] = df[col].iloc[first_valid_loc]
                df.loc[:first_valid_index, col] = df[col].loc[first_valid_index]
            
            if last_valid_index is not None:
                # last_valid_loc = df.index.get_loc(last_valid_index)
                # df[col].iloc[last_valid_loc + 1:] = df[col].iloc[last_valid_loc]
                df.loc[last_valid_index:, col] = df[col].loc[last_valid_index]

        # Sum the columns to create the new totalizer column
        df[new_col] = df[cols].sum(axis=1)
        
        # Create the _actual column
        df[f'{new_col}_actual'] = df[new_col]
        df[f'{new_col}_actual'] = df[f'{new_col}_actual'].ffill()
        
        # Drop summed column
        df = df.drop(columns=new_col)
    
    # Flatten the list of lists and convert to a set to get unique values
    unique_values_totalizer = set(item for sublist in totalizer_dict.values() for item in sublist)

    # Convert the set back to a list if needed
    unique_values_totalizer_list = list(unique_values_totalizer)
    for col in unique_values_totalizer_list:
        # Drop the original columns
        if col in df.columns:
            df = df.drop(columns=col)
        
    return df

def convert_bits(row, columns, unpack_format='d'):
    """Function to convert non-NaN values in the specified columns into a single column."""  
    int_list = [int(row[col]) for col in columns if not np.isnan(row[col])]
    if unpack_format == 'd':
        if len(int_list) != 4:
            return np.nan  # Ensure we have exactly 4 values to process
    elif unpack_format.endswith('I'):
        if len(int_list) != 2:
            return np.nan  # Ensure we have exactly 2 values to process    

    # Pack each 16-bit integer into a 2-byte sequence
    byte_sequence = b''.join(struct.pack('H', n) for n in int_list)
    
    # Interpret the byte sequence based on the specified unpack format
    if unpack_format == 'd':  # Double-precision float
        value = round(struct.unpack('d', byte_sequence)[0], 3)
    elif unpack_format.endswith('I'):  # 32-bit unsigned integer
        value = struct.unpack(unpack_format, byte_sequence)[0]
    else:
        raise ValueError(f"Unsupported unpack format: {unpack_format}")
    
    return value

def convert_to_1_minute_data(df, totalizer_dict):
    """Convert 15-second data to 1-minute data."""
    print("Convert to 1 minute data...")    
    # Set 'Adjusted Timestamp' as the index
    df = df.set_index('Adjusted Timestamp')
    
    # Process the totalizer columns
    df = process_totalizers(df, totalizer_dict)
    
    # Identify non-totalizer columns
    non_totalizer_columns = list(set(df.columns) - set(f'{col}_actual' for col in totalizer_dict.keys()))
    totalizer_columns = [f'{col}_actual' for col in totalizer_dict.keys()]
    
    # Resample to 1-minute intervals and calculate the mean for normal columns
    df_1min_non_totalizer = df[non_totalizer_columns].resample('1min').mean()

    # For totalizer columns, take the value at the zero second within that minute
    df_1min_totalizer = df[totalizer_columns].resample('1min').asfreq()

    # Combine the two DataFrames
    df_1min = pd.concat([df_1min_non_totalizer, df_1min_totalizer], axis=1)

    # Calculate the _diff columns
    for new_col in totalizer_dict.keys():
        df_1min[f'{new_col}_diff'] = df_1min[f'{new_col}_actual'].diff().fillna(0)
    
    # Reset the index to make 'Adjusted Timestamp' a column again
    df_1min = df_1min.reset_index()
    
    return df_1min


def interpolate_nans(df, nLimit=None):
    print('\nInterpolating missing values in df. Current fraction of missing values:')
    dfMissing = pd.DataFrame()
    dfMissing['Sum'] = df.isna().sum()
    dfMissing['Fraction'] = round(df.isna().sum() / df.isna().count(),3)
    print(dfMissing)
# TODO Check if works everywhere
    # df = df.interpolate(method='linear')
    # df = df.interpolate(method='time')
    df = interpolate_columns(df, df.columns, nLimit)
    # And deal with the first row as well:
    # df = df.fillna(method='bfill')
# EDIT python 3.12, pandas 2.2.2
    cols = [col for col in df.columns[3:] if np.isnan(df.loc[df.index[0],col])]
    df.loc[df.index[0], cols] = df.loc[df.index[1], cols]
    # df = df.bfill()
    return df


def write_to_xl(df, dfHeaders, fname=''):
    print('\nWriting output to Excel: ')
    ts = datetime.now().strftime('%Y-%m-%d_%Hh%M')
    if fname=='':
        fname = f'WPdata-{ts}.xlsx'
    else:
        fname = f'WPdata-{fname}.xlsx'
    df1min = df.resample('1T').mean().round(2)
    df1hr = df.resample('1H').mean().round(2)
    
    # Round to 2 decimals first
    # 
    with pd.ExcelWriter(cmb(pExcel,fname)) as writer:
        print('  Hour data..')
        df1hr.to_excel(writer, sheet_name='Hours')
        print('  Minute data..')
        df1min.to_excel(writer, sheet_name='Minutes')
        print('  All data..')
        df.round(2).to_excel(writer, sheet_name='All data')
        print('  Headers..')
        dfHeaders.to_excel(writer, sheet_name='Headers')
        print('Done, closing Excel file.')

def create_output_dataframe(df, lstHeaderMapping):
    # Check if all columns in the header mapping exist in the DataFrame
    print("Create output dataframe...")
    missing_columns = [col for col in lstHeaderMapping if col not in df.columns]
    if missing_columns:
        raise ValueError(f"The following columns are missing in the DataFrame: {missing_columns}")

    # Select the columns and create a new DataFrame
    selected_columns = {col: lstHeaderMapping[col] for col in lstHeaderMapping if col in df.columns}
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

def save_dataframe_with_dates(sMeetsetFolder, df, lstHeaderMapping, folder_name='ExcelOutput'):
    print("Save dataframe to xlsx...")
    df_output = df.copy()
    folder_dir = cmb(folder_name, sMeetsetFolder)
    
    # Ensure the folder exists
    if not os.path.exists(folder_dir):
        os.makedirs(folder_dir)

    # Extract start and end dates from the 'Datum' column
    start_date = df_output[('Datum', '', '')].iloc[0].replace('/', '-')
    end_date = df_output[('Datum', '', '')].iloc[-1].replace('/', '-')

    # Determine the file name based on the dates
    if start_date == end_date:
        file_name = f"Energy Balance - {start_date}.xlsx"
    else:
        file_name = f"Energy Balance - {start_date} - {end_date}.xlsx"

    # Full path to save the file
    file_path = os.path.join(folder_dir, file_name)

    # Get the decimal places mapping
    decimal_places_mapping = get_decimal_places_mapping(lstHeaderMapping)

    # Round the DataFrame columns based on the decimal places mapping
    for col in df_output.columns:
        col_name = col[0]  # Assuming the first level of the MultiIndex is the key
        if col_name in decimal_places_mapping:
            df_output[col] = df_output[col].round(decimal_places_mapping[col_name])

    # Save the DataFrame to an Excel file using pd.ExcelWriter
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df_output.to_excel(writer, index=True, sheet_name='Sheet1')

    print(f"DataFrame saved to {file_path}")

def get_decimal_places_mapping(lstHeaderMapping):
    decimal_places = {}
    
    for key, value in lstHeaderMapping.items():
        mv_key = value[0]
        dimension = value[2]
        if dimension in ['m3/h', 'kWh', 'bara', 'm3(n)/h']:
            decimal_places[mv_key] = 3
        elif dimension in ['\u00b0C', 'hPa']:
            decimal_places[mv_key] = 2
        elif dimension in ['%', 'kW', 'l/h', 'W']:
            decimal_places[mv_key] = 0
        elif dimension == '-':
            # Skip adding to decimal_places mapping
            continue
        else:
            print(f"The key: {key} has not a specificied amount of decimals, please specify it!")
            decimal_places[mv_key] = 2  # Default to 2 decimals if not specified
    return decimal_places


def sortColumns(df, lstStartCols):
    lstToSort = list(set(df.columns) - set(lstStartCols))
    lstToSort.sort()
    lstNewOrder = lstStartCols + lstToSort
    df = df[lstNewOrder]
    return df


if __name__ == "__main__":
    
    bReadData = False
    bWriteExcel = False
    bDebugStopExecutionHere = True
    
    sMeetsetFolder = 'StresstestDeventer_08-12jul'
    # sMeetsetFolder = 'Deventer_13-17jul'
    # sMeetsetFolder = 'Meetset2-Nunspeet'
    # sMeetsetFolder = 'Meetset1-Deventer'

    # Set environment variables
    pBase = os.getcwd()
    pInput = cmb(pBase,'Data')
    pPickles = cmb(pBase,'ImportedPickles')
    pExcel = cmb(pBase,'ExcelOutput')
    # Read data into pickles
    if bReadData:
        process_and_save(pBase, sMeetsetFolder)

    # Load stored data from pickles, process and write to Excel
    sDate = '2024-07-01'
    sDateEnd = '2025-12-31'
    dfRaw = load_data(sMeetset=sMeetsetFolder, sDateStart = sDate, sDateEnd = sDateEnd)
    df, dfHeaders = flatten_data(dfRaw, bStatus=False, ignore_multi_index_differences=True)
    # raise ValueError("Stopping here for debugging..")

    # BACKUP columns for if Weather Air Temp value is off, 
    # then reconstruct with a separate function from 03A+03B most likely
    weatherAirTempCols = [col for col in df.columns if col.startswith('Weather Temp Air')]
    df.drop(weatherAirTempCols[1:], axis=1, inplace=True)
    if 'Itron Gas volume 2' in df.columns:
        df.drop(['Itron Gas volume 2','Itron Gas volume 3'], axis=1, inplace=True)

    process_stream_1p = ['Stream1 PressureA', 'Stream1 PressureB', 'Stream1 PressureC', 'Stream1 PressureD']
    process_stream_1t = ['Stream1 TemperatureA', 'Stream1 TemperatureB', 'Stream1 TemperatureC', 'Stream1 TemperatureD']
    process_stream_1f = ['Stream1 FlowA', 'Stream1 FlowB', 'Stream1 FlowC', 'Stream1 FlowD']
    process_stream_2p = ['Stream2 PressureA', 'Stream2 PressureB', 'Stream2 PressureC', 'Stream2 PressureD']
    process_stream_2t = ['Stream2 TemperatureA', 'Stream2 TemperatureB', 'Stream2 TemperatureC', 'Stream2 TemperatureD']
    process_stream_2f = ['Stream2 FlowA', 'Stream2 FlowB', 'Stream2 FlowC', 'Stream2 FlowD']
    
    print("Convert bits to values...")
    # Apply the function to each row and create a new column with the results
    dictBitConversionEVHI = {'Stream1 Pressure':process_stream_1p, 'Stream1 Temperature':process_stream_1t, 
                             'Stream1 Flow':process_stream_1f, 'Stream2 Pressure':process_stream_2p, 
                             'Stream2 Temperature':process_stream_2t,'Stream2 Flow':process_stream_2f}
    for key,value in dictBitConversionEVHI.items():
        if value[0] in df.columns:
            df[key] = df.apply(lambda row: convert_bits(row, value, unpack_format='d'), axis=1)
            df.drop(value, axis=1, inplace=True)

    # process_belimo_1f = ['Belimo01 FlowTotalM3A', 'Belimo01 FlowTotalM3B']
    # process_belimo_1h = ['Belimo01 Heating EnergyA', 'Belimo01 Heating EnergyB']
    # process_belimo_2f = ['Belimo02 FlowTotalM3A', 'Belimo02 FlowTotalM3B']
    # process_belimo_2h = ['Belimo02 Heating EnergyA', 'Belimo02 Heating EnergyB']
    # process_belimo_3f = ['Belimo03 FlowTotalM3A', 'Belimo03 FlowTotalM3B']
    # process_belimo_3h = ['Belimo03 Heating EnergyA', 'Belimo03 Heating EnergyB']
    # process_belimo_vf = ['BelimoValve FlowTotalM3A', 'BelimoValve FlowTotalM3B']
    # process_belimo_vh = ['BelimoValve Heating EnergyA', 'BelimoValve Heating EnergyB']
    # process_belimo_1frate = ['Belimo01 FlowRateA', 'Belimo01 FlowRateB']
    # df['Belimo01 FlowTotalM3'] = df.apply(lambda row: convert_bits(row, process_belimo_1f, unpack_format='I'), axis=1)
    # df['Belimo01 Heating Energy'] = df.apply(lambda row: convert_bits(row, process_belimo_1h, unpack_format='I'), axis=1)
    # df['Belimo02 FlowTotalM3'] = df.apply(lambda row: convert_bits(row, process_belimo_2f, unpack_format='I'), axis=1)
    # df['Belimo02 Heating Energy'] = df.apply(lambda row: convert_bits(row, process_belimo_2h, unpack_format='I'), axis=1)
    # df['Belimo03 FlowTotalM3'] = df.apply(lambda row: convert_bits(row, process_belimo_3f, unpack_format='I'), axis=1)
    # df['Belimo03 Heating Energy'] = df.apply(lambda row: convert_bits(row, process_belimo_3h, unpack_format='I'), axis=1)
    # df['BelimoValve FlowTotalM3'] = df.apply(lambda row: convert_bits(row, process_belimo_vf, unpack_format='I'), axis=1)
    # df['BelimoValve Heating Energy'] = df.apply(lambda row: convert_bits(row, process_belimo_vh, unpack_format='I'), axis=1)
    # df['Belimo01 FlowRate'] = df.apply(lambda row: convert_bits(row, process_belimo_1frate, unpack_format='I'), axis=1)
    # colsBelimoDrop = process_belimo_1f + process_belimo_1h + process_belimo_2f + process_belimo_2h \
    #                 + process_belimo_3f + process_belimo_3h + process_belimo_vf + process_belimo_vh + process_belimo_1frate
    # df.drop(colsBelimoDrop, axis=1, inplace=True)

    if bDebugStopExecutionHere:    
        raise ValueError("Stopping here for debugging..")
    else:
        df = combine_and_sync_rows(df)
        df = add_hours_based_on_dst(df)

    
    # Exctract the keys from header list that are needed for this code to run
    # Remove diff if added, to get the original columns
    columns_keys = list(hHP.makeAllHeaderMappings().keys())
    original_columns = [item.replace('_diff', '') if item.endswith('_diff') else item for item in columns_keys]
    # Some columns are created, ignore these
    ignore_columns = ['Contains missing data', 'Missing data (no Eastron02)', 
                      'Missing data (no Belimo)', 'Eastron Total Power']
    needed_columns = list(set(original_columns)- set(ignore_columns))
    # Check that only existing columns remain
    needed_columns = list(set(needed_columns).intersection(set(df.columns)))
    # Some columns are a lot of times NaN, include those
    ignore_eastron2_column = [col for col in needed_columns if col.startswith('Eastron02')]
    needed_columns_without_eastron2 = list(set(needed_columns)- set(ignore_eastron2_column))
    exclude_belimo_columns = [col for col in needed_columns if col.startswith('Belimo')]
    needed_columns_without_belimo = list(set(needed_columns_without_eastron2)- set(exclude_belimo_columns))
    # df['Contains missing data'] = df[needed_columns].isnull().any(axis=1)
    df['Missing data (no Eastron02)'] = df[needed_columns_without_eastron2].isnull().any(axis=1)
    df['Missing data (no Belimo)'] = df[needed_columns_without_belimo].isnull().any(axis=1)

    if 'Weather Abs Air Pressure' in df.columns:
        # Convert pgasin from barg to bara, but if air pressure is below 700 mbar, add 1.023 to the value instead of reading the air pressure
        df['PgasIn'] = np.where(
            df['Weather Abs Air Pressure'] < 700, 1.023 + df['PgasIn'],
            df['Weather Abs Air Pressure'].div(1000) + df['PgasIn']
        )
    if 'Eastron01 Total Power' in df.columns and 'Eastron02 Total Power' in df.columns:
        df['Eastron Total Power'] = df[['Eastron01 Total Power', 'Eastron02 Total Power']].sum(axis=1)
    df = sortColumns(df, ['Adjusted Timestamp','Missing data (no Eastron02)','Missing data (no Belimo)'])
    # Write to Excel the full dataframe
    # print('Saving full dataframe...')
    # df.to_excel(cmb(pExcel, sMeetsetFolder + datetime.now().strftime('_%Y-%m-%d_%Hh%M.xlsx')))
    df.set_index('Adjusted Timestamp', drop=False, inplace=True)
    df = interpolate_nans(df, nLimit=10)

#TODO Check on totalizer columns that they are monotonic increasing and if not, interpolate so it works out.
    # BelimoValve FlowTotalL
#DONE Check on outlier values:
    # TgasIn: Outliers for temperature
    rowsMask = df[~df['TgasIn'].between(-5.0,100.0)].index
    df.loc[rowsMask, 'TgasIn'] = np.nan
    rowsMask = df[~df['PgasIn'].between(20.0,70.0)].index
    df.loc[rowsMask, 'PgasIn'] = np.nan
    
    
    #Now work on the 1min output for the energy balance calculation
    dict_totalizer = {
        'Itron Gas volume 1': ['Itron Gas volume 1'],
        'Belimo01 FlowTotalL': ['Belimo01 FlowTotalL'],
        'Belimo02 FlowTotalL': ['Belimo02 FlowTotalL'],
        'Belimo03 FlowTotalL': ['Belimo03 FlowTotalL'],
        'BelimoValve FlowTotalL': ['BelimoValve FlowTotalL'],
        }
    dict_totalizer = {k: dict_totalizer[k] for k in dict_totalizer if k in df.columns}
    df_1min = convert_to_1_minute_data(df, dict_totalizer)
    df_1min.set_index('Adjusted Timestamp', drop=False, inplace=True)
    #Remove lines where no data was logged, so check a few basic columns are all nan:
    dropIndices = df_1min[df_1min[['Missing data (no Belimo)','PgasIn','TgasIn']].isna().all(axis=1)].index        
    
    df_1min.drop(dropIndices.unique(), axis=0, inplace=True)

    
    if 'Itron Gas volume 1_diff' in df_1min.columns:
        df_1min['Itron Gas volume 1_diff'] = df_1min['Itron Gas volume 1_diff']*60*1000 # Convert first from m3*0.1/minute to m3*0.1/h and then to l/h
    if 'Belimo03 FlowTotalL_diff' in df_1min.columns:
        df_1min['Belimo01 FlowTotalL_diff'] = df_1min['Belimo01 FlowTotalL_diff']*60 # Convert from litre/minute to to l/h
        df_1min['Belimo02 FlowTotalL_diff'] = df_1min['Belimo02 FlowTotalL_diff']*60 # Convert from litre/minute to to l/h
        df_1min['Belimo03 FlowTotalL_diff'] = df_1min['Belimo03 FlowTotalL_diff']*60 # Convert from litre/minute to to l/h
        df_1min['BelimoValve FlowTotalL_diff'] = df_1min['BelimoValve FlowTotalL_diff']*60 # Convert from litre/minute to to l/h

    #DONE HERE: copy flow rates where _diff exists and 'FlowRate' is empty, to FlowRate
    if bDebugStopExecutionHere:
        belimoCorr = ['01','02','03','Valve']
        for sNum in belimoCorr:
            rowsNoFlow =  df_1min[df_1min[f'Belimo{sNum} FlowRate'].isna()].index
            rowsWithData = df_1min[df_1min[[f'Belimo{sNum} Temp1 external',f'Belimo{sNum} Temp2 internal',f'Belimo{sNum} FlowTotalL_diff']].notna().all(axis=1)].index
            rowsMask = rowsNoFlow.intersection(rowsWithData)
            df_1min.loc[rowsMask, f'Belimo{sNum} FlowRate'] = df_1min.loc[rowsMask,f'Belimo{sNum} FlowTotalL_diff']

                
    # Use .sum because that treats the nans in Eastron02 column as zeros.
    #Wrong, was this ever used? 
    # df_1min['Eastron01 Total Power'] = df_1min['Eastron01 Total Power']*60 # Convert from kWh/minute to kW
    # df_1min['Eastron02 Total Power'] = df_1min['Eastron02 Total Power']*60 # Convert from kWh/minute to kW
    
    # lstBelimoTotals = ['Belimo01 FlowTotalM3', 'Belimo02 FlowTotalM3', 'Belimo03 FlowTotalM3']
    # for col in lstBelimoTotals:
    #     if col in df_1min.columns:
    #         df_1min[col] = df_1min[col]*60*10 # Convert first from m3*100/minute to m3/h and than to l/h
    df_1min = sortColumns(df_1min, ['Adjusted Timestamp','Missing data (no Eastron02)','Missing data (no Belimo)'])
    df_1min = interpolate_nans(df_1min, nLimit=10)

    if bDebugStopExecutionHere:
        df_1min.to_excel(cmb(pExcel, '1min_' + sMeetsetFolder + datetime.now().strftime('_%Y-%m-%d_%Hh%M.xlsx')))

    if bWriteExcel:
        # lstHeaderMapping = hHP.makeAllHeaderMappings()
        dictHeaderMapping = hHP.genHeaders(df_1min.columns)
        df_1min_newheaders = create_output_dataframe(df_1min, dictHeaderMapping)
        save_dataframe_with_dates(sMeetsetFolder, df_1min_newheaders, dictHeaderMapping)
    