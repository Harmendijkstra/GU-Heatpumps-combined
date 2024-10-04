# -*- coding: utf-8 -*-
"""
Created on Mon Jul 29 16:48:27 2024

@author: LENLUI
"""

import win32com.client, pythoncom
# import shutil
# from os.path import join as cmb
import os

xlApp, xlWb, xlSheet = (None,)*3


def xlOpen(bXLvisible = 0):
    print("Opening Excel Application...")
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    # Optional to use Ensuredispatch to launch Excel in different manner so that code completion works different
    xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
    #xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
    xlApp.Visible = bXLvisible
    return xlApp


def xlClose(xl):
    print("Closing Excel Application...")
    xl.DisplayAlerts = False
#    wDoc.Close()
    xl.Application.Quit()
    xl = None
    del xl
#    pythoncom.CoUninitialize()

def openWorkbook(xl, filepath):
    if os.path.isfile(filepath):
        xlWb = xl.Workbooks.Open(filepath)
    else:
        print(f'File not found: {filepath}')
        return None
    return xlWb

def closeWorkbook(xl, xlWb, bSave = False):
    if bSave:
        xlWb.Save()
    xl.DisplayAlerts = False
    xlWb.Close()
    
def autofitColumns(xlSheet):
    xlSheet.Columns.AutoFit()
    
def removeRow(xlSheet, nRow):
    xlSheet.Rows(nRow).Delete()



    
