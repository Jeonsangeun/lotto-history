# __author__ = "Ashwin Nanjappa"

# GUI viewer to view JSON data as tree.
# Ubuntu packages needed:
# python3-pyqt5

# Std
import collections
import json
import sys

# External
from PyQt5 import QtCore
from PyQt5 import QtWidgets


class TextToTreeItem:

    def __init__(self):
        self.text_list = []
        self.titem_list = []

    def append(self, text_list, titem):
        for text in text_list:
            self.text_list.append(text)
            self.titem_list.append(titem)

    # Return model indices that match string
    def find(self, find_str):

        titem_list = []
        for i, s in enumerate(self.text_list):
            if find_str in s:
                titem_list.append(self.titem_list[i])

        return titem_list


class JsonView(QtWidgets.QWidget):

    def __init__(self):
        super(JsonView, self).__init__()

        self.tree_widget = None
        self.text_to_titem = TextToTreeItem()

        # jdata = json.load(jfile, object_pairs_hook=collections.OrderedDict)
        self.jdata = {}

        # Find UI
        copyright_layout = self.copyright_ui()
        choice_layout = self.choose_ui()

        # Tree format UI

        self.tree_widget = QtWidgets.QTreeWidget()
        self.tree_widget.setHeaderLabels(["Key", "Value"])
        self.tree_widget.header().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        root_item = QtWidgets.QTreeWidgetItem(["Root"])
        self.tree_widget.addTopLevelItem(root_item)
        # Add table to layout

        layout = QtWidgets.QHBoxLayout()
        layout.addWidget(self.tree_widget)

        # Group box

        self.gbox = QtWidgets.QGroupBox('No select')
        self.gbox.setLayout(layout)

        self.layout2 = QtWidgets.QVBoxLayout()
        self.layout2.addLayout(copyright_layout)
        self.layout2.addLayout(choice_layout)
        self.layout2.addWidget(self.gbox)

        self.setLayout(self.layout2)

    def copyright_ui(delf):
        creator = QtWidgets.QLabel("Made by Ashwin Nanjappa")
        editor = QtWidgets.QLabel("Edit by Sang-eun Jeon")

        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(creator)
        layout.addWidget(editor)

        return layout

    def choose_ui(self):
        search_button = QtWidgets.QPushButton("Path")
        clear_button = QtWidgets.QPushButton("Clear")

        search_button.clicked.connect(self.search_button_clicked)
        clear_button.clicked.connect(self.clear_button_clicked)

        layout = QtWidgets.QHBoxLayout()
        layout.addWidget(search_button)
        layout.addWidget(clear_button)

        return layout

    def Build_json_data(self, path):
        # re-build json data
        self.layout2.removeWidget(self.gbox)
        if path == 'None':
            self.jdata = {}
        else:
            dict = open(path)
            self.jdata = json.load(dict, object_pairs_hook=collections.OrderedDict)

        self.tree_widget = QtWidgets.QTreeWidget()
        self.tree_widget.setHeaderLabels(["Key", "Value"])
        self.tree_widget.header().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        root_item = QtWidgets.QTreeWidgetItem(["Root"])
        self.recurse_jdata(self.jdata, root_item)
        self.tree_widget.addTopLevelItem(root_item)

        layout = QtWidgets.QHBoxLayout()
        layout.addWidget(self.tree_widget)

        self.gbox = QtWidgets.QGroupBox(path)
        self.gbox.setLayout(layout)

        self.layout2.addWidget(self.gbox)
        self.setLayout(self.layout2)

    def search_button_clicked(self):
        f_path = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File', '', 'JSON(*.json)')
        if f_path[0][-5:] != '.json':
            QtWidgets.QMessageBox.about(self, 'Alert', 'Please select .json')
        else:
            # re-build json data
            self.Build_json_data(f_path[0])

    def clear_button_clicked(self):
        self.Build_json_data('None')

    def recurse_jdata(self, jdata, tree_widget):

        if isinstance(jdata, dict):
            for key, val in jdata.items():
                self.tree_add_row(key, val, tree_widget)
        elif isinstance(jdata, list):
            for i, val in enumerate(jdata):
                key = str(i)
                self.tree_add_row(key, val, tree_widget)
        else:
            print("This should never be reached!")

    def tree_add_row(self, key, val, tree_widget):

        text_list = []

        if isinstance(val, dict) or isinstance(val, list):
            text_list.append(key)
            row_item = QtWidgets.QTreeWidgetItem([key])
            row_item.setFlags(row_item.flags() | QtCore.Qt.ItemIsEditable)
            self.recurse_jdata(val, row_item)
        else:
            text_list.append(key)
            text_list.append(str(val))
            row_item = QtWidgets.QTreeWidgetItem([key, str(val)])
            row_item.setFlags(row_item.flags() | QtCore.Qt.ItemIsEditable)

        tree_widget.addChild(row_item)
        self.text_to_titem.append(text_list, row_item)


class JsonViewer(QtWidgets.QMainWindow):

    def __init__(self):
        super(JsonViewer, self).__init__()
        json_view = JsonView()

        self.setCentralWidget(json_view)
        self.setWindowTitle("JSON Viewer")
        self.show()

    def keyPressEvent(self, e):
        if e.key() == QtCore.Qt.Key_Escape:
            self.close()

qt_app = QtWidgets.QApplication(sys.argv)
json_viewer = JsonViewer()
sys.exit(qt_app.exec_())