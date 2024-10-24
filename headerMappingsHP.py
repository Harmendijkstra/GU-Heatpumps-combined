# -*- coding: utf-8 -*-
"""
Created on Thu Jun 24 10:07:49 2021

@author: LENLUI
"""


def genHeaders(lstSingleHeader):
    dictMapping = makeAllHeaderMappings()
    # lstHeaders = [lstMapping[x] for x in lstSingleHeader]
    dictHeaders = {}
    for key,value in dictMapping.items():
        if key in lstSingleHeader:
            dictHeaders.update({key: dictMapping[key]})
        # else:
        #     dictHeaders.update({key: ('','','')})
    # for h in lstSingleHeader:
    #     if h in dictMapping:
    #         dictHeaders.update({h: dictMapping[h]})
    return dictHeaders


def makeAllHeaderMappings():
    lstHeaderMapping = {
        # 'Contains missing data':        ['MV0' ,'Contains missing data','-'],
        'Missing data (no Eastron02)':        ['Missing1' ,'Missing data (no Eastron02)','-'],
        'Missing data (no Belimo)':        ['Missing2' ,'Missing data (no Belimo)','-'],
        'Weather Temp Air':             ['MV1' ,'T_omg','\u00b0C'],
        'Weather Abs Air Pressure':     ['MV2' ,'p_baro','hPa'],
        'Weather Rel Humidity':         ['MV3' ,'RV','%'],
        'PgasIn':                       ['MV4' ,'p_gas_in','bara'],
        'TgasIn':                       ['MV5' ,'T_gas_in','\u00b0C'],
        'Stream1 Flow':                 ['MV6' ,'V_gas_str1','m3(n)/h'],
        'Stream1 Pressure':             ['MV7' ,'p_uit_str1','bara'],
        'Stream1 Temperature':          ['MV8' ,'T_uit_str1','\u00b0C'],
        'Stream2 Flow':                 ['MV9' ,'V_gas_str2','m3(n)/h'],
        'Stream2 Pressure':             ['MV10' ,'p_uit_str2','bara'],
        'Stream2 Temperature':          ['MV11' ,'T_uit_str2','\u00b0C'],
        'Itron Gas volume 1_diff':      ['MV15','V_gas_br','l/h'],
        'Eastron01 Total Power':   ['MV16a','Pe_WP1','W'],
        'Eastron02 Total Power':   ['MV16b','Pe_WP2','W'],
        'Eastron Total Power':     ['MV16','Pe_WP','W'],
        # 'Belimo01 FlowTotalM3':         ['MV17','Vw_ket_1','l/h'],
        # 'Belimo02 FlowTotalM3':         ['MV20','Vw_ex_OV','l/h'],
        # 'Belimo03 FlowTotalM3':         ['MV21','Vw_WP','l/h'],
        'Belimo01 FlowRate':         ['MV17','Vw_ket_1','l/h'],
        'Belimo02 FlowRate':         ['MV20','Vw_ex_OV','l/h'],
        'Belimo03 FlowRate':         ['MV21','Vw_WP','l/h'],
        'Belimo01 Temp1 external':      ['MV22','Tw_ket1_in','\u00b0C'], 
        'Belimo01 Temp2 internal':      ['MV23','Tw_ket1_uit','\u00b0C'],
        'Belimo02 Temp2 internal':      ['MV28','Ts_OV_uit','\u00b0C'],
        'Belimo02 Temp1 external':      ['MV29','Ts_OV_in','\u00b0C'],
        'ADAM PT1000 01':               ['MV30','Tb_1','\u00b0C'],
        'ADAM PT1000 02':               ['MV31','Tb_2','\u00b0C'],
        'ADAM PT1000 03':               ['MV32','Tb_3','\u00b0C'],
        'ADAM PT1000 04':               ['MV33','Tb_4','\u00b0C'],
        'Belimo03 Temp1 external':      ['MV34','T_WP_in','\u00b0C'],
        'Belimo03 Temp2 internal':      ['MV35','T_WP_uit','\u00b0C'],
        # 'BelimoValve FlowRate':            ['MV36','Vw_klep','l/h'],
        # 'BelimoValve Temp1 external':      ['MV37','Tw_klep_in','\u00b0C'], 
        # 'BelimoValve Temp2 internal':      ['MV38','Tw_klep_uit','\u00b0C'],
        'Q_ket1_wm':                {"minute_data": ['AV1','Q_ket1_wm','kJ/s'], "hourly_data": ['AV1','Q_ket1_wm','kWh']}, #Note that this is a dictonary to have different dimesions for minute and hourly data
        'Q_OV_wm':                  {"minute_data": ['AV2','Q_OV_wm','kJ/s'], "hourly_data": ['AV2','Q_OV_wm','kWh']}, #Note that this is a dictonary to have different dimesions for minute and hourly data
        'Q_WP_wm':                  {"minute_data": ['AV3','Q_WP_wm','kJ/s'], "hourly_data": ['AV3','Q_WP_wm','kWh']},#Note that this is a dictonary to have different dimesions for minute and hourly data
        # 'Q_klep_wm_sum':                ['AV4','Q_klep_wm_sum','kJ/s'], #Note that this is a dictonary to have different dimesions for minute and hourly data
        'Cop_fabr1':                      ['AV5','Cop_fabr1','-'],
        'Cop_fabr2':                      ['AV6','Cop_fabr2','-'],
    }
    return lstHeaderMapping





