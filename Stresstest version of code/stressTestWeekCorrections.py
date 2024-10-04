# -*- coding: utf-8 -*-
"""
Created on Mon Jul 22 12:13:36 2024

@author: LENLUI
"""

# Code to manually edit the Stress Test week data, to produce a valid result.
# Save the code in use here, so it can be reproduced in case it's needed.

# Stop the code manually above the Belimo processing lines and above combine/sync.


colsDrop = [col for col in df.columns if col.startswith('Belimo') and (col.endswith('Custom') \
           or col.endswith('Heating Energy'))]
colsDrop = colsDrop + ['Stream1 Pressure1'] + [col for col in df.columns if col.endswith('Import')]

# df['Itron Gas volume 1'][df['Itron Gas volume 1']> 1000000] = 0.0
df.loc[df['Itron Gas volume 1']> 1000000, 'Itron Gas volume 1'] = 0.0

belimoCorr = ['01','02','03','Valve']
for sNum in belimoCorr:
    sCol = f'Belimo{sNum} FlowTotal m3'
    destCols = [f'Belimo{sNum} FlowTotalM3A', f'Belimo{sNum} FlowTotalM3B']
    mask = df[df[sCol]>0].index
    print(f"Correcting {sCol} now...")
    for rowLabel in mask:
        df.loc[rowLabel, destCols] = [struct.unpack(">HH", struct.pack('>I', int(df.loc[rowLabel,sCol]*100) ) )[0], 0]
    colsDrop = colsDrop + [sCol]
        
print('Converting bits to values - Belimo')
process_belimo_1f = ['Belimo01 FlowTotalM3A', 'Belimo01 FlowTotalM3B']
process_belimo_1h = ['Belimo01 Heating EnergyA', 'Belimo01 Heating EnergyB']
process_belimo_2f = ['Belimo02 FlowTotalM3A', 'Belimo02 FlowTotalM3B']
process_belimo_2h = ['Belimo02 Heating EnergyA', 'Belimo02 Heating EnergyB']
process_belimo_3f = ['Belimo03 FlowTotalM3A', 'Belimo03 FlowTotalM3B']
process_belimo_3h = ['Belimo03 Heating EnergyA', 'Belimo03 Heating EnergyB']
process_belimo_vf = ['BelimoValve FlowTotalM3A', 'BelimoValve FlowTotalM3B']
process_belimo_vh = ['BelimoValve Heating EnergyA', 'BelimoValve Heating EnergyB']
process_belimo_1frate = ['Belimo01 FlowRateA', 'Belimo01 FlowRateB']
df['Belimo01 FlowTotalM3'] = df.apply(lambda row: convert_bits(row, process_belimo_1f, unpack_format='I'), axis=1)
df['Belimo01 Heating Energy'] = df.apply(lambda row: convert_bits(row, process_belimo_1h, unpack_format='I'), axis=1)
df['Belimo02 FlowTotalM3'] = df.apply(lambda row: convert_bits(row, process_belimo_2f, unpack_format='I'), axis=1)
df['Belimo02 Heating Energy'] = df.apply(lambda row: convert_bits(row, process_belimo_2h, unpack_format='I'), axis=1)
df['Belimo03 FlowTotalM3'] = df.apply(lambda row: convert_bits(row, process_belimo_3f, unpack_format='I'), axis=1)
df['Belimo03 Heating Energy'] = df.apply(lambda row: convert_bits(row, process_belimo_3h, unpack_format='I'), axis=1)
df['BelimoValve FlowTotalM3'] = df.apply(lambda row: convert_bits(row, process_belimo_vf, unpack_format='I'), axis=1)
df['BelimoValve Heating Energy'] = df.apply(lambda row: convert_bits(row, process_belimo_vh, unpack_format='I'), axis=1)
df['Belimo01 FlowRate 1000L'] = df.apply(lambda row: convert_bits(row, process_belimo_1frate, unpack_format='I'), axis=1)
colsBelimoDrop = process_belimo_1f + process_belimo_1h + process_belimo_2f + process_belimo_2h \
                + process_belimo_3f + process_belimo_3h + process_belimo_vf + process_belimo_vh + process_belimo_1frate
df.drop(colsBelimoDrop, axis=1, inplace=True)

print('Dropping erroneous data, correcting Belimo totalizers')
dropIndices = df[df['Belimo01 FlowTotalM3']<1E4].index
dropIndices = dropIndices.append(df[df['Belimo02 FlowTotalM3']<1E3].index)
dropIndices = dropIndices.append(df[df['Belimo03 FlowTotalM3']<1E3].index)
dropIndices = dropIndices.append(df[df['BelimoValve FlowTotalM3']<10].index)
dropIndices = dropIndices.append(df.loc[df['Belimo01 FlowTotalM3']>1E8,'Belimo01 FlowTotalM3'].index)
df.drop(dropIndices.unique(), axis=0, inplace=True)

df = combine_and_sync_rows(df)
df = add_hours_based_on_dst(df)

belimoCorr = ['01','02','03','Valve']
nEndRowX10 = df[df['Belimo01 FlowTotalL']>0].index[0]
for sNum in belimoCorr:
    sCol = f'Belimo{sNum} FlowTotalM3'
    maskRows = df[df[sCol].between(1,1E5)].index
    # .loc[df.index[0]:nEndRowX10
    if maskRows[-1] > 20969:
        print('Might be wrong here')
    df.loc[maskRows, f'Belimo{sNum} FlowTotalL'] = 10 * df.loc[maskRows,sCol]
    if sNum=='01':
        maskRows = df[df[sCol] > 1E5].index
        df.loc[maskRows, f'Belimo{sNum} FlowTotalL'] = df.loc[maskRows,sCol]
    # df.loc[df.index[0]:nEndRowX10, f'Belimo{sNum} FlowTotalL'] = 10 * df.loc[df.index[0]:nEndRowX10,sCol]
    # df[f'Belimo{sNum} FlowTotalL'] = df[[f'Belimo{sNum} FlowTotalL',sCol]].sum(axis=1)
    colsDrop = colsDrop + [sCol]

df['Belimo01 FlowRate 1000L'] = df['Belimo01 FlowRate 1000L'] / 1000
maskRows = df[df['Belimo01 FlowRate 1000L'] > 0].index
df.loc[maskRows, 'Belimo01 FlowRate'] = df.loc[maskRows, 'Belimo01 FlowRate 1000L']
colsDrop = colsDrop + ['Belimo01 FlowRate 1000L']

maskRows = df[df['Itron Gas volume 1']>100.0].index
df.loc[maskRows, 'Itron Gas volume 1'] = df.loc[maskRows, 'Itron Gas volume 1'] / 10
firstValidTotalizerValue = df[df['Itron Gas volume 1']> 0.1]['Itron Gas volume 1'].index[0]
df.loc[df.index[0]:firstValidTotalizerValue, 'Itron Gas volume 1'] = df.loc[firstValidTotalizerValue, 'Itron Gas volume 1']

df.drop(colsDrop, axis=1, inplace=True)    
#Correct one BelimoValve error
df.drop(df.index[59], axis=0, inplace=True)

###
# Last part, when all corrections are done manually, quick save with below code.
###


if 'Timestamp' in df.columns:
    df = sortColumns(df,['Timestamp'])
elif 'Adjusted Timestamp' in df.columns:
    df = sortColumns(df,['Adjusted Timestamp'])
print('Writing full dataframe to file')    
df.to_excel(cmb(pExcel, 'RawData_' + sMeetsetFolder + datetime.now().strftime('_%Y-%m-%d_%Hh%M.xlsx')))