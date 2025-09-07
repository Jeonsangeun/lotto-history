import sys
import pandas as pd

from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QTextEdit, QComboBox, QLabel, QMessageBox, QHBoxLayout, QVBoxLayout
from PyQt5.QtCore import QCoreApplication

class MyApp(QWidget):

    def __init__(self):
        super().__init__()

        self.meta_data = pd.read_excel('POST_CODE.xlsx', sheet_name=None, na_values="TT")
        self.name = QLabel('None', self)
        self.post = QLineEdit('AA', self)
        self.post.setMaxLength(2)
        self.S_des = QTextEdit('According to the manufacturer', self)
        self.D_des = QTextEdit('According to the DCT experience', self)
        self.note = QLineEdit('note', self)

        self.initUI()

    def initUI(self):

        cb = QComboBox(self)
        for item in ['Wiwynn', 'Quanta', "Ingrasys", "Lenovo", "DELL", "HPE"]:
            cb.addItem(item)
        cb.activated[str].connect(self.setting_name)

        M_Button = QPushButton('Search!', self)
        S_Button = QPushButton('save', self)
        cancelButton = QPushButton('Exit', self)

        cancelButton.clicked.connect(QCoreApplication.instance().quit)
        M_Button.clicked.connect(self.search)
        M_Button.clicked.connect(self.DCT_search)

        hbox0 = QHBoxLayout()
        hbox0.addWidget(QLabel('Manufacturer'))
        hbox0.addWidget(cb)
        hbox0.addWidget(self.name)
        hbox0.addStretch(1)

        hbox1 = QHBoxLayout()
        hbox1.addWidget(QLabel('POST CODE'))
        hbox1.addWidget(self.post)
        hbox1.addWidget(M_Button)
        hbox1.addStretch(1)

        hbox2 = QHBoxLayout()
        hbox2.addWidget(QLabel('Specification'))
        hbox2.addStretch(1)

        hbox3 = QHBoxLayout()
        hbox3.addWidget(self.S_des)
        hbox3.addStretch(1)

        hbox4 = QHBoxLayout()
        hbox4.addWidget(QLabel('DCT Comments'))
        hbox4.addStretch(1)

        hbox5 = QHBoxLayout()
        hbox5.addWidget(self.D_des)
        hbox5.addStretch(1)

        hbox6 = QHBoxLayout()
        hbox6.addWidget(QLabel('Note'))
        hbox6.addWidget(self.note)
        hbox6.addWidget(S_Button)
        hbox6.addStretch(1)

        hbox7 = QHBoxLayout()
        hbox7.addWidget(cancelButton)
        hbox7.addStretch(1)

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(hbox0)
        vbox.addStretch(1)
        vbox.addLayout(hbox1)
        vbox.addStretch(1)
        vbox.addLayout(hbox2)
        vbox.addStretch(1)
        vbox.addLayout(hbox3)
        vbox.addStretch(1)
        vbox.addLayout(hbox4)
        vbox.addStretch(1)
        vbox.addLayout(hbox5)
        vbox.addStretch(1)
        vbox.addLayout(hbox6)
        vbox.addStretch(1)
        vbox.addLayout(hbox7)
        vbox.addStretch(1)

        self.setLayout(vbox)

        self.setWindowTitle('POST Library')
        self.setGeometry(100, 100, 250, 500)
        self.show()

    def setting_name(self, text):
        self.name.setText(text)

    def search(self):
        if self.name.text() == 'None':
            QMessageBox.about(self, 'Alert', 'Please select manufacturer')
        else:
            sheet = self.name.text()
            data = self.meta_data[sheet]
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
            data = self.meta_data['DCT']
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


app = QApplication(sys.argv)
ex = MyApp()
sys.exit(app.exec_())