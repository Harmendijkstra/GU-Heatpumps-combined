# -*- coding: utf-8 -*-
"""
Created on Wed Jun 19 11:14:24 2024

@author: LENLUI
"""

import numpy as np
import pandas as pd
from io import StringIO

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
    for row in data_rows:
        if row.startswith('D'):
            parts = row.split(';')
            timestamp = parts[1]
            data_dict = {sTimestampCol: timestamp}
            for i in range(2, len(parts), 3):
                if i + 2 < len(parts):
                    signal_number = int(parts[i])
                    if parts[i + 1] == '':
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

if __name__ == "__main__":
    fPath = 'sample-OMC-045_DNV_Deventer_048000232_240618_210000.txt'
    sData, header_df = read_file(fPath)
    dfEmpty = create_multi_index_df(header_df)
    df = process_datafile(sData, dfEmpty, header_df)
    
    # Display the DataFrame
    print(df)
    print("\nOptionally, separate the df into a 'Value' and a 'Status' DataFrame.")
    print("Then, on the values, interpolate to the lowest 15sec multiple")