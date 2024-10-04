# -*- coding: utf-8 -*-
"""
Created on Wed Jun 19 11:14:24 2024

@author: LENLUI
"""

import pandas as pd
from io import StringIO

# Sample data
data = """T;25;26;27;28;29;30;31;32;33;34;35;36;37;38;39;40;41;42;43;44;45;46
Timestamp;Itron Gas volume 1;Itron Gas volume 2;Itron Gas volume 3;Hager Voltage 1;Weather Temp Air;Weather Temp Housing;Weather Rel Humidity;Weather Abs Humidity;Weather Abs Air Pressure;Belimo01 Temp1 external;Belimo01 Temp2 internal;Belimo01 FlowTotal;Belimo01 FlowTotal;Belimo01 Heating Energy;ADAM PT1000 01;ADAM PT1000 02;ADAM PT1000 03;ADAM PT1000 04;ADAM PT1000 05;ADAM PT1000 06;TgasIn;PgasIn
yy/mm/dd hh:mm:ss;m3;m3;m3;V;C;C;%;g/m3;hPa;C;C;m3;l/min?;m3;C;C;C;C;C;C;C;mA
Time;ItronVol1_IntesisMbus_modbus_all;ItronVol2_IntesisMbus_modbus_all;ItronVol3_IntesisMbus_modbus_all;HagerV1_IntesisMbus_modbus_all;WTempAir_ThiesWeather_modbus_all;WTempHouse_ThiesWeather_modbus_all;WHumRel_ThiesWeather_modbus_all;WHumAbs_ThiesWeather_modbus_all;WHumAbs_ThiesWeather_modbus_all;Belimo01tExt_Belimo01_modbus_all;Belimo01tInt_Belimo01_modbus_all;Belimo01flow_Belimo01_modbus_all;Belimo01flow_Belimo01_modbus_all;Belimo01HeatE_Belimo01_modbus_all;ADAMt01_Adam4015_modbus_all;ADAMt02_Adam4015_modbus_all;ADAMt03_Adam4015_modbus_all;ADAMt04_Adam4015_modbus_all;ADAMt05_Adam4015_modbus_all;ADAMt06_Adam4015_modbus_all;AC1;AC2
D;24/06/18 21:00:00;25;0;0;26;0;0;27;0;0;28;239;0;29;22.7;0;30;34.0;0;31;51.8;0;32;10.4;0;33;1014.6;0;34;21.18;0;35;21.0;0;36;0.0;0;37;0;0;38;0;0;39;159.9997;0;40;159.9997;0;41;159.9997;0;42;159.9997;0;43;159.9997;0;44;159.9997;0
D;24/06/18 21:00:12;46;4.287128;0;45;20.90104;0
D;24/06/18 21:00:15;25;0;0;26;0;0;27;0;0;28;239;0;29;22.7;0;30;34.0;0;31;51.8;0;32;10.4;0;33;1014.6;0;34;21.18;0;35;21.0;0;36;0.0;0;37;0;0;38;0;0;39;159.9997;0;40;159.9997;0;41;159.9997;0;42;159.9997;0;43;159.9997;0;44;159.9997;0
D;24/06/18 21:00:27;46;4.287128;0;45;20.90104;0
D;24/06/18 21:00:30;25;0;0;26;0;0;27;0;0;28;239;0;29;22.6;0;30;34.0;0;31;52.10001;0;32;10.41;0;33;1014.5;0;34;21.18;0;35;20.99;0;36;0.0;0;37;0;0;38;0;0;39;159.9997;0;40;159.9997;0;41;159.9997;0;42;159.9997;0;43;159.9997;0;44;159.9997;0
D;24/06/18 21:00:42;46;4.287128;0;45;20.91937;0
D;24/06/18 21:00:45;25;0;0;26;0;0;27;0;0;28;239;0;29;22.6;0;30;34.0;0;31;52.10001;0;32;10.41;0;33;1014.5;0;34;21.17;0;35;20.99;0;36;0.0;0;37;0;0;38;0;0;39;159.9997;0;40;159.9997;0;41;159.9997;0;42;159.9997;0;43;159.9997;0;44;159.9997;0
D;24/06/18 21:00:57;46;4.287128;0;45;20.91937;0
D;24/06/18 21:01:00;25;0;0;26;0;0;27;0;0;28;239;0;29;22.6;0;30;34.0;0;31;52.0;0;32;10.39;0;33;1014.6;0;34;21.18;0;35;21.0;0;36;0.0;0;37;0;0;38;0;0;39;159.9997;0;40;159.9997;0;41;159.9997;0;42;159.9997;0;43;159.9997;0;44;159.9997;0
D;24/06/18 21:01:12;45;20.90104;0;46;4.287128;0
"""

# Read the header rows
header_data = StringIO('\n'.join(data.split('\n')[:4]))
header_df = pd.read_csv(header_data, sep=';', header=None)

# Create the multi-index
nSignals = len(header_df.iloc[0])
# multi_index = pd.MultiIndex.from_arrays([header_df.iloc[1], header_df.iloc[2], header_df.iloc[3], header_df.iloc[0]], names=['Signal Name', 'Unit', 'Tag', 'Signal No.'])
multi_indexValues = pd.MultiIndex.from_arrays([header_df.iloc[1], header_df.iloc[2], header_df.iloc[3], header_df.iloc[0], ['Value']*nSignals], names=['Signal Name', 'Unit', 'Tag', 'Signal No.', 'Type'])
multi_indexStatus = pd.MultiIndex.from_arrays([header_df.iloc[1, 1:], header_df.iloc[2, 1:], header_df.iloc[3, 1:], header_df.iloc[0, 1:], ['Status']*(nSignals-1)], names=['Signal Name', 'Unit', 'Tag', 'Signal No.', 'Type'])
lst = multi_indexValues.to_flat_index().append(multi_indexStatus.to_flat_index())
multi_index = pd.MultiIndex.from_tuples(lst, names=['Signal Name', 'Unit', 'Tag', 'Signal No.', 'Type'])

# Initialize an empty DataFrame with the multi-index columns
df = pd.DataFrame(columns=multi_index)
sTimestampCol = df.columns[0]

# Initialize a list to collect rows
rows = []

# Read the data rows
data_rows = data.split('\n')[4:]

# Process each data row
for row in data_rows:
    if row.startswith('D'):
        parts = row.split(';')
        timestamp = parts[1]
        data_dict = {sTimestampCol: timestamp}
        for i in range(2, len(parts), 3):
            if i + 2 < len(parts):
                signal_number = int(parts[i])
                signal_value = float(parts[i + 1])
                status_bit = int(parts[i + 2])
                signal_name = header_df.iloc[1, signal_number - 24]
                unit = header_df.iloc[2, signal_number - 24]
                tag = header_df.iloc[3, signal_number - 24]
                data_dict[(signal_name, unit, tag, signal_number, 'Value')] = signal_value
                data_dict[(signal_name, unit, tag, signal_number, 'Status')] = status_bit
        # df = df.append(data_dict, ignore_index=True)
        rows.append(data_dict)

# Create the DataFrame from the list of rows
df = pd.DataFrame(rows)
        
# Set the multi-index columns
df.columns = pd.MultiIndex.from_tuples(df.columns, names=['Signal Name', 'Unit', 'Tag', 'Signal No.', 'Type'])

# Display the DataFrame
print(df)