import sys
import pandas as pd
import openpyxl as excel

from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QHBoxLayout, QVBoxLayout, QLabel, QComboBox, QFileDialog, QMessageBox

Available_CELL = ["F01C01", "F01C02", "F01C03", "F01C04", "F02C01", "F02C02", "F02C03", "F02C04", "F03C01", "F03C02", "F03C03", "F03C04"]

def isNaN(num):
    return num != num

class MyApp(QWidget):

    def __init__(self):
        super().__init__()

        self.Colo_select = QLabel('Select Colo : ', self)
        self.PRD_res_url = QLabel('', self)
        self.res_mas_url = QLabel('', self)
        self.UPS_apply = QLabel('현재 적용된 COLO : ', self)
        self.UPS_rack = {'UPS1': [], 'UPS2': [], 'UPS3': [], 'UPS4': []}
        self.COLO_log = [0] * len(Available_CELL)

        self.initUI()

    def initUI(self):

        cb = QComboBox(self)
        for item in Available_CELL:
            cb.addItem(item)
        cb.activated[str].connect(self.setting_name)

        res_url = QPushButton("PRD reservation URL")
        res_url.clicked.connect(self.url_clicked)

        UPS_select = QPushButton("Find UPS")
        UPS_select.clicked.connect(self.UPS_find)

        master_url = QPushButton("Master list URL")
        master_url.clicked.connect(self.url_for_mas)

        Divide_UPS1 = QPushButton("UPS1")
        Divide_UPS1.clicked.connect(self.function_transfer_1)

        Divide_UPS2 = QPushButton("UPS2")
        Divide_UPS2.clicked.connect(self.function_transfer_2)

        Divide_UPS3 = QPushButton("UPS3")
        Divide_UPS3.clicked.connect(self.function_transfer_3)

        Divide_UPS4 = QPushButton("UPS4")
        Divide_UPS4.clicked.connect(self.function_transfer_4)

        hbox = QHBoxLayout()
        hbox.addWidget(res_url)
        hbox.addStretch(1)
        hbox.addWidget(self.PRD_res_url)
        hbox.addStretch(1)

        hbox1 = QHBoxLayout()
        hbox1.addWidget(cb)
        hbox1.addWidget(UPS_select)
        hbox1.addWidget(self.Colo_select)
        hbox1.addStretch(1)

        hbox2 = QHBoxLayout()
        hbox2.addStretch(1)
        hbox2.addWidget(self.UPS_apply)
        hbox2.addStretch(1)

        hbox3 = QHBoxLayout()
        hbox3.addWidget(master_url)
        hbox3.addStretch(1)
        hbox3.addWidget(self.res_mas_url)
        hbox3.addStretch(1)

        hbox4 = QHBoxLayout()
        hbox4.addWidget(Divide_UPS1)
        hbox4.addStretch(1)
        hbox4.addWidget(Divide_UPS2)
        hbox4.addStretch(1)
        hbox4.addWidget(Divide_UPS3)
        hbox4.addStretch(1)
        hbox4.addWidget(Divide_UPS4)
        hbox4.addStretch(1)

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(hbox)
        vbox.addStretch(1)
        vbox.addLayout(hbox1)
        vbox.addStretch(1)
        vbox.addLayout(hbox2)
        vbox.addStretch(1)
        vbox.addLayout(hbox3)
        vbox.addStretch(1)
        vbox.addLayout(hbox4)
        vbox.addStretch(1)

        self.setLayout(vbox)

        self.setWindowTitle('Box Layout')
        self.setGeometry(300, 500, 500, 200)
        self.show()

    def setting_name(self, text):
        self.Colo_select.setText(text)

    def url_clicked(self):
        f_path = QFileDialog.getOpenFileName(self, 'Open File', '', 'EXCEL(*.xlsx)')
        if f_path[0][-5:] != '.xlsx':
            QMessageBox.about(self, 'Alert', 'Please select .xlsx')
        else:
            self.PRD_res_url.setText(f_path[0])

    def url_for_mas(self):
        f_path = QFileDialog.getOpenFileName(self, 'Open File', '', 'EXCEL(*.xlsx)')
        if f_path[0][-5:] != '.xlsx':
            QMessageBox.about(self, 'Alert', 'Please select .xlsx')
        else:
            self.res_mas_url.setText(f_path[0])

    def UPS_find(self):
        if self.Colo_select.text() == 'Select Colo : ' or self.PRD_res_url.text() == '':
            QMessageBox.about(self, 'Alert', 'Please select Colo or URL')
        elif self.COLO_log[Available_CELL.index(self.Colo_select.text())] == 1:
            QMessageBox.about(self, 'Caution', 'Already tired it')
        else:
            load_df = pd.read_excel(self.PRD_res_url.text(), sheet_name=self.Colo_select.text())
            UPS = ['UPS1', 'UPS2', 'UPS3', 'UPS4']

            if 'DOOR' in list(load_df.keys()):
                flag, RNG = 0, ""
                for key, value in load_df.items():
                    if flag == 1 and value[0][0] != 'U':
                        RNG = value[0]
                        flag = 0
                    if len(str(value[1])) == 2 and RNG != 'RNG':
                        for text in value:
                            if text in UPS:
                                self.UPS_rack[text].append((self.Colo_select.text(), str(value[1])))
                    if not isNaN(value[0]):
                        power_source = value[0][0]
                        if power_source.isdecimal():
                            flag = 1
            else:
                flag, RNG = 0, ""
                for key, value in load_df.items():
                    if flag == 1 and key[0] != 'U':
                        flag, RNG = 0, key
                    if len(str(value[0])) == 2 and RNG != 'RNG':
                        for text in value:
                            if text in UPS:
                                self.UPS_rack[text].append((self.Colo_select.text(), str(value[0])))
                    power_source = key[0]  # find xxxKW
                    if power_source.isdecimal():
                        flag = 1

            self.COLO_log[Available_CELL.index(self.Colo_select.text())] = 1
            self.UPS_apply.setText(self.UPS_apply.text() + self.Colo_select.text() + ' ,')

    def function_transfer_1(self):
        if self.UPS_rack == {'UPS1': [], 'UPS2': [], 'UPS3': [], 'UPS4': []}:
            QMessageBox.about(self, 'Alert', 'No data applied.')
        else:
            return self.Divided_EG("UPS1")

    def function_transfer_2(self):
        if self.UPS_rack == {'UPS1': [], 'UPS2': [], 'UPS3': [], 'UPS4': []}:
            QMessageBox.about(self, 'Alert', 'No data applied.')
        else:
            return self.Divided_EG("UPS2")

    def function_transfer_3(self):
        if self.UPS_rack == {'UPS1': [], 'UPS2': [], 'UPS3': [], 'UPS4': []}:
            QMessageBox.about(self, 'Alert', 'No data applied.')
        else:
            return self.Divided_EG("UPS3")

    def function_transfer_4(self):
        if self.UPS_rack == {'UPS1': [], 'UPS2': [], 'UPS3': [], 'UPS4': []}:
            QMessageBox.about(self, 'Alert', 'No data applied.')
        else:
            return self.Divided_EG("UPS4")

    def Divided_EG(self, ups):

        if self.res_mas_url.text() == '':
            QMessageBox.about(self, 'Alert', 'Please select URL')
        else:
            wb = excel.Workbook()

            load_ex = pd.read_excel(self.res_mas_url.text())
            EG = []
            Location = []

            title = list(load_ex.keys())
            index_DcL = list(load_ex.keys()).index('DCLocation')
            index_SN = list(load_ex.keys()).index('ServiceName')

            for data in load_ex['ServiceName']:
                if not isNaN(data):
                    EG.append(data)
            EG_set = list(set(EG))

            for loc in self.UPS_rack[ups]:
                Location.append(loc[0] + '.' + loc[1])

            for value in EG_set:
                wb.create_sheet(value[:30], index=0)
                ws = wb[value[:30]]
                ws.append(title)
                for data_res in load_ex.values:
                    if not isNaN(data_res[index_DcL]):
                        if data_res[index_DcL][6:15] in Location:
                            if data_res[index_SN] == value:
                                ws.append(list(data_res))

            wb.save("PUS04-XXXX-DD-MM-YYYY IT Power Maintenance-" + ups + ".xlsx")
            QMessageBox.about(self, 'Success', 'Complete to save')

app = QApplication(sys.argv)
ex = MyApp()
sys.exit(app.exec_())
