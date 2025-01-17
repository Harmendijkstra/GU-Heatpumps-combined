import pandas as pd
import numpy as np
from scipy.interpolate import interp2d
from os.path import join as cmb
from createOutput import convert_excel_output, save_dataframe_with_dates
import os

def interpolate_2d(df_Ggas, temperature, pressure):
    # Find the surrounding temperatures and pressures
    temp_cols = df_Ggas.columns
    temp_idx = np.searchsorted(temp_cols, temperature)
    if temp_idx == 0:
        temp_idx = 1
    T1 = temp_cols[temp_idx - 1]
    T2 = temp_cols[temp_idx]

    pressure_rows = df_Ggas.index
    pressure_idx = np.searchsorted(pressure_rows, pressure)
    if pressure_idx == 0:
        pressure_idx = 1
    p1 = pressure_rows[pressure_idx - 1]
    p2 = pressure_rows[pressure_idx]

    # Calculate the relative positions
    relT = (temperature - T1) / (T2 - T1)
    relp = (pressure - p1) / (p2 - p1)

    # Extract the surrounding values
    z1 = df_Ggas.loc[p1, T1]
    z2 = df_Ggas.loc[p2, T1]
    z3 = df_Ggas.loc[p1, T2]
    z4 = df_Ggas.loc[p2, T2]

    # Perform the interpolation
    zp1 = z1 + relp * (z2 - z1)
    zp2 = z3 + relp * (z4 - z3)
    return zp1 + relT * (zp2 - zp1)

def load_enthalpy_table(file_path):
    dfEnthalpy = pd.read_excel(file_path, skiprows=[0, 1])
    df_Ggas = dfEnthalpy.iloc[:17, 1:]
    df_Ggas = df_Ggas.set_index('Unnamed: 1')
    return df_Ggas

def create_interpolation_function(df_Ggas):
    temperatures = df_Ggas.columns[0:].astype(float)
    pressures = df_Ggas.index
    values = df_Ggas.values
    interp_func = interp2d(temperatures, pressures, values, kind='linear')
    return interp_func, temperatures, pressures

def interpolate_h_gas_in(row, temp_col, pres_col, interp_func, temperatures, pressures):
    temperature = row[temp_col]
    pressure = row[pres_col]
    
    if (temperature < temperatures.min() or temperature > temperatures.max() or
        pressure < pressures.min() or pressure > pressures.max()):
        np.nan
    
    return interp_func(temperature, pressure)[0]

def perform_calculations(df_Ggas, df, interp_func, temperatures, pressures):
    def calculate_and_compare(row, temp_col, pres_col):
        temperature = row[temp_col]
        pressure = row[pres_col]
        
        if pd.isnull(temperature) or pd.isnull(pressure):
            return np.nan
        
        library_value = interpolate_2d(df_Ggas, temperature, pressure)
        own_function_value = interpolate_h_gas_in(
            row, 
            temp_col=temp_col, 
            pres_col=pres_col,
            interp_func=interp_func,
            temperatures=temperatures,
            pressures=pressures
        )
        
        if np.round(library_value, 2) != np.round(own_function_value, 2):
            print(f"Warning: different values at temperature {temperature} and pressure {pressure}: library={library_value}, own_function={own_function_value}")
        
        return library_value
    
    df[('RV1', 'h_gas_in', 'kJ/kg')] = df.apply(
        lambda row: calculate_and_compare(row, ('MV5', 'T_gas_in', '\u00b0C'), ('MV4', 'p_gas_in', 'bara')),
        axis=1
    )
    
    df[('RV2', 'h_uit_str1', 'kJ/kg')] = df.apply(
        lambda row: calculate_and_compare(row, ('MV8', 'T_uit_str1', '\u00b0C'), ('MV7', 'p_uit_str1', 'bara')),
        axis=1
    )
    
    df[('RV3', 'm_gas_str1', 'kg/h')] = df[('MV6','V_gas_str1', 'm3(n)/h')] *  0.8334
    df[('RV4', 'dh_str1', 'kJ/kg')] = (df[('RV2', 'h_uit_str1', 'kJ/kg')].sub(df[('RV1', 'h_gas_in', 'kJ/kg')])).clip(lower=0)
    df[('RV5', 'Q_str1', 'kJ/s')] =  df[('RV3', 'm_gas_str1', 'kg/h')].mul(df[('RV4', 'dh_str1', 'kJ/kg')]) / 3600
    df[('RV6', 'm_gas_str2', 'kg/h')] = df[('MV9', 'V_gas_str2', 'm3(n)/h')] *  0.8334

    df[('RV7', 'h_uit_str2', 'kJ/kg')] = df.apply(
        lambda row: calculate_and_compare(row, ('MV11', 'T_uit_str2', '\u00b0C'), ('MV10', 'p_uit_str2', 'bara')),
        axis=1
    )

    df[('RV8', 'dh_str2', 'kJ/kg')] = (df[('RV7', 'h_uit_str2', 'kJ/kg')].sub(df[('RV1', 'h_gas_in', 'kJ/kg')])).clip(lower=0)
    df[('RV9', 'Q_str2', 'kJ/s')] = df[('RV6', 'm_gas_str2', 'kg/h')].mul(df[('RV8', 'dh_str2', 'kJ/kg')]) / 3600
    df[('RV10', 'Q_brgas', 'kJ/s')] = ((df[('MV15', 'V_gas_br', 'l/h')] /1000) * 35.17) / 3.6
    
    # dT_ketelw = df[('MV23', 'Tw_ket1_uit', '\u00b0C')].sub(df[('MV22', 'Tw_ket1_in', '\u00b0C')])
    # df[('RV11', 'Q_ket1', 'kJ/s')] = (df[('MV17', 'Vw_ket_1', 'l/h')] * 4.19 * (dT_ketelw)) / 3600
    dT_ketelw = df[('MV28', 'Ts_OV_uit', '\u00b0C')].sub(df[('MV29', 'Ts_OV_in', '\u00b0C')])
    df[('RV11', 'Q_ket1', 'kJ/s')] = (df[('MV20', 'Vw_ex_OV', 'l/h')] * 4.19 * (dT_ketelw)) / 3600 #Note that this is changed, using the OV instead of the ketel

    dT_WP = df[('MV35', 'T_WP_uit', '\u00b0C')].sub(df[('MV34', 'T_WP_in', '\u00b0C')])
    df[('RV14', 'Q_WP', 'kJ/s')] = (df[('MV21', 'Vw_WP', 'l/h')] * 4.19 * (dT_WP)) / 3600
    
    # dT_koeler = df[('MV38', 'Tw_klep_uit', '\u00b0C')].sub(df[('MV37', 'Tw_klep_in', '\u00b0C')])
    # df[('RV15', 'Q_koeler', 'kJ/s')] = (df[('MV36', 'Vw_klep', 'l/h')] * 4.19 * (dT_koeler)) / 3600
    dT_koeler = 0  # We no longer use this, but this is still here such that the same code can be used if above is no longer commented
    df[('RV16', 'Rend_ket', '%')] = (df[('RV11', 'Q_ket1', 'kJ/s')]).div(df[('RV10', 'Q_brgas', 'kJ/s')].where(df[('RV10', 'Q_brgas', 'kJ/s')] != 0, np.nan)) * 100
    
    Pe_WP_kW = df[('MV16', 'Pe_WP', 'W')]/1000
    df[('RV17', 'Rend_WP', '%')] = df[('RV14', 'Q_WP', 'kJ/s')].div(Pe_WP_kW.where(Pe_WP_kW != 0, np.nan)) * 100
    
    Q_afgeg = df[('RV11', 'Q_ket1', 'kJ/s')].add(df[('RV14', 'Q_WP', 'kJ/s')])
    Q_opgen = df[('RV10', 'Q_brgas', 'kJ/s')].add(Pe_WP_kW)
    df[('RV18', 'Rend_waterz', '%')] = Q_afgeg.div(Q_opgen.where(Q_opgen != 0, np.nan)) * 100
    
    # Q_straten = df[('RV5', 'Q_str1', 'kJ/s')].add(df[('RV9', 'Q_str2', 'kJ/s')]).add(df[('RV15', 'Q_koeler', 'kJ/s')])
    Q_straten = df[('RV5', 'Q_str1', 'kJ/s')].add(df[('RV9', 'Q_str2', 'kJ/s')])
    df[('RV19', 'Rend_tot', '%')] = Q_straten.div(Q_opgen.where(Q_opgen != 0, np.nan)) * 100
    return df, dT_ketelw, dT_WP, dT_koeler, Pe_WP_kW, Q_afgeg, Q_opgen, Q_straten

def add_additional_columns(df, dT_ketelw, dT_WP, dT_koeler, Pe_WP_kW, Q_afgeg, Q_opgen, Q_straten):
    df_full = df.copy()
    df_full[('RV11a', 'dT_ketelw', '')] = dT_ketelw
    df_full[('RV14a', 'dT_WP', '')] = dT_WP
    # df_full[('RV15a', 'dT_koeler', '')] = dT_koeler
    df_full[('RV16a', 'Pe_WP_kW', '')] = Pe_WP_kW
    df_full[('RV18a', 'Q_afgeg', '')] = Q_afgeg
    df_full[('RV18b', 'Q_opgen', '')] = Q_opgen
    df_full[('RV19a', 'Q_straten', '')] = Q_straten
    return df_full

def save_and_convert(df, prefix, weekFolder):
    fpath = save_dataframe_with_dates(df, {}, folder_dir=weekFolder, prefix=prefix, header_input_type='df')
    change_files = [fpath]
    convert_excel_output(weekFolder, change_files)
    return fpath

def add_enthalpy_calcualations(df_1hr_newheaders, folder_dir, year, prefix=''):
    sEnthalpyTable = 'EnthalpyInput/EnthalpyTable.xlsx'    
    df_Ggas = load_enthalpy_table(sEnthalpyTable)
    interp_func, temperatures, pressures = create_interpolation_function(df_Ggas)
    
    df_1hr_newheaders_withRV = df_1hr_newheaders.copy()
    
    df_1hr_newheaders_withRV, dT_ketelw, dT_WP, dT_koeler, Pe_WP_kW, Q_afgeg, Q_opgen, Q_straten = perform_calculations(df_Ggas, df_1hr_newheaders_withRV, interp_func, temperatures, pressures)
    
    df_full = add_additional_columns(df_1hr_newheaders_withRV, dT_ketelw, dT_WP, dT_koeler, Pe_WP_kW, Q_afgeg, Q_opgen, Q_straten)
    
    # Get the unique weeks from the DataFrame index
    iso_calendar = df_full.index.isocalendar()
    all_weeks = list(set(zip(iso_calendar.year, iso_calendar.week)))
    all_weeks.sort() # Sort the weeks
    weeks_with_year = []

    for year, no in all_weeks:
        mask = (iso_calendar.year == year) & (iso_calendar.week == no)
        this_week = df_full[mask]
        
        # Check whether the week is complete
        if mask.sum() == 0:
            print(f'Skipping week {no} of year {year}, no data')
            continue
        time_diff = this_week.index[-1] - this_week.index[0]
        skipping = False
        if not time_diff >= pd.Timedelta('6 days 23:00:00'):
            skipping = True

        if skipping:
            print(f'Skipping week {no} of year {year}, not complete week')
        else:
            print(f'Saving week no {no} of year {year}')
            df_1hr_newheaders_withRV_week = df_1hr_newheaders_withRV[mask]
            df_1hr_newheaders_week = df_1hr_newheaders[mask]

            # Get the first date in the DataFrame
            first_date = df_1hr_newheaders_week.index[0]
            # Get the ISO year and week number
            iso_year, iso_week, _ = first_date.isocalendar()

            df_full_week = df_full[mask]
            weekFolder = os.path.join(str(folder_dir), f"{iso_year}-{iso_week:02d}")
            weeks_with_year.append(weekFolder)
            if not os.path.exists(weekFolder):
                os.makedirs(weekFolder)
            weekPrefixRV = f"{prefix}weekno - {iso_week}"
            weekPrefix = f"{prefix.replace('RV - ', '')}weekno - {iso_week}"
            save_and_convert(df_1hr_newheaders_withRV_week, weekPrefixRV, weekFolder)
            save_and_convert(df_1hr_newheaders_week, weekPrefix, weekFolder)
            # fullweekPrefix = weekPrefix + 'full - '
            # save_and_convert(df_full_week, fullweekPrefix, weekFolder)
    return weeks_with_year