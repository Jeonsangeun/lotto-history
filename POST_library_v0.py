import sys
import pandas as pd
import datetime as dt
import os
import time

from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QTextEdit, QComboBox, QLabel, QGridLayout, QMessageBox
from PyQt5.QtCore import QCoreApplication

if getattr(sys, 'frozen', False):
    # test.exe로 실행한 경우,test.exe를 보관한 디렉토리의 full path를 취득
    program_directory = os.path.dirname(os.path.abspath(sys.executable))
else:
    # python test.py로 실행한 경우,test.py를 보관한 디렉토리의 full path를 취득
    program_directory = os.path.dirname(os.path.abspath(__file__))

class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        grid = QGridLayout()
        self.setLayout(grid)
        cb = QComboBox(self)
        for item in ['Wiwynn', 'Quanta', "Ingrasys", "Lenovo", "DELL", "HPE"]:
            cb.addItem(item)
        self.name = QLabel('None', self)
        self.post = QLineEdit('AA', self)
        self.M_Button = QPushButton('Search!', self)
        self.S_des = QTextEdit('According to the manufacturer', self)
        self.D_des = QTextEdit('According to the DCT experience', self)
        self.note = QLineEdit('note', self)
        self.S_Button = QPushButton('save', self)

        self.post.setMaxLength(2)

        cancelButton = QPushButton('Exit', self)
        cancelButton.clicked.connect(QCoreApplication.instance().quit)

        self.M_Button.clicked.connect(self.search)
        self.M_Button.clicked.connect(self.DCT_search)

        grid.addWidget(QLabel('Manufacturer'), 0, 0)
        grid.addWidget(QLabel('POST CODE'), 2, 0)
        grid.addWidget(QLabel('Specification'), 4, 0)
        grid.addWidget(QLabel('DCT Comments'), 6, 0)
        grid.addWidget(QLabel('Note'), 8, 0)

        grid.addWidget(cb, 1, 0)
        grid.addWidget(self.name, 1, 1)
        cb.activated[str].connect(self.setting_name)

        grid.addWidget(self.post, 3, 0)
        grid.addWidget(self.M_Button, 3, 1)

        grid.addWidget(self.S_des, 5, 0)
        grid.addWidget(self.D_des, 7, 0)
        grid.addWidget(self.note, 9, 0)
        grid.addWidget(self.S_Button, 9, 1)
        grid.addWidget(cancelButton, 10, 0)

        self.setWindowTitle('POST Library')
        self.setGeometry(300, 300, 500, 300)
        self.show()

    def setting_name(self, text):
        self.name.setText(text)

    def search(self):
        if self.name.text() == 'None':
            QMessageBox.about(self, 'Alert', 'Please select manufacturer')
        else:
            sheet = self.name.text()
            data = pd.read_excel('POST_CODE.xlsx', sheet_name=sheet, na_values="TT")
            data = data.fillna('')
            find_code = self.post.text().upper()
            description = ''
            for value in data.values:
                if value[0] == find_code:
                    description += value[1]
                    description += '\n'
            if description == '':
                self.S_des.setText('No found')
            else:
                self.S_des.setText(description)

    def DCT_search(self):
        relate_task, note_history, date, DC_code = 3, 4, 6, 7
        if self.name.text() == 'None':
            pass
        else:
            sheet = 'DCT'
            data = pd.read_excel('POST_CODE.xlsx', sheet_name=sheet, na_values='=')
            data = data.fillna('')
            find_code = self.post.text().upper()
            note = ''
            for value in data.values:
                if value[0] == find_code:
                    if value[date] != '':
                        time_date = value[date]
                        note += ('GDCO : ' + value[relate_task][-10:] + " | ")
                        note += ('date : ' + time_date.strftime('%Y-%m-%d') + " | ")
                        note += ('DC : ' + value[DC_code])
                        note += '\n'
                        note += ('note : ' + value[note_history])
                        note += '\n'
                    else:
                        pass
            if note == '':
                self.D_des.setText('No found')
            else:
                self.D_des.setText(note)

starttime = time.time()

app = QApplication(sys.argv)
ex = MyApp()

print("time : ", time.time() - starttime)
sys.exit(app.exec_())