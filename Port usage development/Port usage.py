import sys
import openpyxl as excel
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
import os

from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QHBoxLayout, QVBoxLayout, QTextEdit, QLabel, QFileDialog
from PyQt5.QtCore import QCoreApplication

# load_df = pd.read_excel('test.xlsx')

load_df = pd.read_csv('sample.csv', index_col=0)
temp_data = dict(load_df)
usage_device = ['C0', 'M0']

class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.url = QLabel('url', self)
        self.error_message = QTextEdit('', self)

        self.initUI()

    def initUI(self):
        F_Button = QPushButton('Search', self)
        F_Button.clicked.connect(self.search_button_clicked)

        PRD_Button = QPushButton('Go!', self)
        PRD_Button.clicked.connect(self.PRD)

        hbox = QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(F_Button)
        hbox.addStretch(1)
        hbox.addWidget(self.url)
        hbox.addStretch(1)

        hbox1 = QHBoxLayout()
        hbox1.addWidget(self.error_message)

        hbox2 = QHBoxLayout()
        hbox2.addWidget(PRD_Button)

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(hbox)
        vbox.addStretch(1)
        vbox.addLayout(hbox1)
        vbox.addStretch(1)
        vbox.addLayout(hbox2)
        vbox.addStretch(1)

        self.setLayout(vbox)

        self.setWindowTitle('Box Layout')
        self.setGeometry(300, 300, 300, 200)
        self.show()

    def search_button_clicked(self):
        f_path = QFileDialog.getOpenFileName(self, 'Open File', '')
        self.url.setText(f_path[0])

    def PRD(self):
        default = "Final_Result"
        load_file_name = self.url.text()
        data = pd.read_excel(load_file_name, sheet_name=default, na_values='=')

        col_Device = list(data.keys()).index('EndDevice')
        col_port = list(data.keys()).index('EndPort')
        description = ""

        for idx, value in enumerate(data.values):
            if value[col_Device][-2:] in usage_device:
                Rack = value[col_port - 1][:5]
                Device = value[col_Device]
                Port = value[col_port]
                if Rack not in temp_data.keys():
                    temp_data[Rack] = {Device: [Port]}
                else:
                    if Device not in temp_data[Rack].keys():
                        temp_data[Rack][Device] = [Port]
                    else:
                        if Port not in temp_data[Rack][Device]:
                            temp_data[Rack][Device].append(Port)
                        else:
                            description += "해당포트에 중복이 있습니다. {} 장비의 {} 포트가 중복입니다.".format(Device, Port)
                            description += '\n'
        if description == '':
            self.error_message.setPlainText('No found')
        else:
            self.error_message.setPlainText(description)

        df = pd.DataFrame(temp_data)
        df.to_csv('sample.csv')

app = QApplication(sys.argv)
ex = MyApp()
sys.exit(app.exec_())

