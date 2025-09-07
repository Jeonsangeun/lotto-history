import os
import sys
import pandas as pd
import openpyxl as excel

from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QHBoxLayout, QVBoxLayout, QLineEdit, QLabel, QComboBox, QFileDialog, QMessageBox
from PyQt5.QtCore import QCoreApplication

Available_CELL = ["F01C01", "F01C02", "F01C03", "F01C04", "F02C01", "F02C02", "F02C03", "F02C04", "F03C01", "F03C02", "F03C03", "F03C04"]
UPS = ['UPS1', 'UPS2', 'UPS3', 'UPS4']
UPS_rack = {'UPS1': [], 'UPS2': [], 'UPS3': [], 'UPS4': []}
load_df = pd.read_excel("PUS04-COLO PRD Reservation v2.0.xlsx", sheet_name=Available_CELL[0])

def isNaN(num):
    return num != num

# available col detect # 0 34 35
flag, RNG = 0, ""
for key, value in load_df.items():
    if flag == 1 and key[0] != 'U':
        flag, RNG = 0, key
        print(RNG)
    if len(str(value[0])) == 2 and RNG != 'RNG':
        for text in value:
            if text in UPS:
                UPS_rack[text].append(("F01C01", str(value[0])))
    power_source = key[0]  # find xxxKW
    if power_source.isdecimal():
        flag = 1

print(UPS_rack)

wb = excel.Workbook()
# wb.active.title = 'Sheet1'
# w1 = wb['Sheet1']

load_ex = pd.read_excel("Resource Master List.xlsx")
EG = []

title = list(load_ex.keys())
index_DcL = list(load_ex.keys()).index('DCLocation')
index_SN = list(load_ex.keys()).index('ServiceName')

for data in load_ex['ServiceName']:
    if not isNaN(data):
        EG.append(data)

EG_set = list(set(EG))
Location = []
for loc in UPS_rack['UPS1']:
    Location.append(loc[0] + '.' + loc[1])
print(Location)

for value in EG_set:
    wb.create_sheet(value[:30], index=0)
    ws = wb[value[:30]]
    ws.append(title)
    for data_res in load_ex.values:
        if not isNaN(data_res[index_DcL]):
            if data_res[index_DcL][6:15] in Location:
                if data_res[index_SN] == value:
                    ws.append(list(data_res))

wb.save("PUS04-XXXX-DD-MM-YYYY IT Power Maintenance-" + 'dfd' + ".xlsx")