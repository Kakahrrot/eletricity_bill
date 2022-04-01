import sys
import pandas
from PyQt5.QtCore import QAbstractTableModel, Qt
from PyQt5.QtWidgets import (QWidget, QLabel, QLineEdit, QCheckBox, QPushButton,
                             QTextEdit, QVBoxLayout, QHBoxLayout, QMainWindow, QApplication,
                             QTableView)
from electricity import electricity, init_elec


class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, index):
        return self._data.shape[0]

    def columnCount(self, parnet=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole or role == Qt.EditRole:
                value = self._data.iloc[index.row(), index.column()]
                return str(value)

    def setData(self, index, value, role):
        if role == Qt.EditRole:
            self._data.iloc[index.row(), index.column()] = value
            window.setData(index, value)
        if role == Qt.EditRole:
            return True
        return False

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]

    def flags(self, index):
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable


class MainWindow(QMainWindow):
	def __init__(self):
		super().__init__()
		self.initUI()
		
		self.data = []

	def initUI(self):
		h0 = QHBoxLayout()
		self.path = QLineEdit()
		self.YearNum = QLineEdit()
		h0.addWidget(QLabel("Path"))
		h0.addWidget(self.path)
		h0.addWidget(QLabel("YearNum"))
		h0.addWidget(self.YearNum)

		v1 = QVBoxLayout()
		v1.addWidget(QLabel("file"))
		self.file = QLineEdit()
		v1.addWidget(self.file)

		v2 = QVBoxLayout()
		v2.addWidget(QLabel("col"))
		self.col = QLineEdit()
		v2.addWidget(self.col)

		v3 = QVBoxLayout()
		v3.addWidget(QLabel("col_name"))
		self.col_name = QLineEdit()
		v3.addWidget(self.col_name)

		v4 = QVBoxLayout()
		v4.addWidget(QLabel("sheet_name"))
		self.sheet_name = QLineEdit()
		v4.addWidget(self.sheet_name)

		v5 = QVBoxLayout()
		self.Center = QCheckBox("Center")
		v5.addWidget(self.Center)
		self.ADD = QPushButton("ADD")
		self.ADD.pressed.connect(self.addOperation)
		v5.addWidget(self.ADD)
		
		self.table = QTableView()
		
		self.Run = QPushButton("RUN")
		self.Run.pressed.connect(self.runOperation)

		h1 = QHBoxLayout()
		h1.addLayout(v1)
		h1.addLayout(v2)
		h1.addLayout(v3)
		h1.addLayout(v4)
		h1.addLayout(v5)

		vv = QVBoxLayout()
		vv.addLayout(h0)
		vv.addLayout(h1)
		vv.addWidget(self.table)
		vv.addWidget(self.Run)
		widget = QWidget()
		widget.setLayout(vv)
		self.setCentralWidget(widget)
		self.setWindowTitle('電表計算')

	def addOperation(self):
		# print(self.file.text(), self.col.text(), self.col_name.text(), self.sheet_name.text(), self.Center.isChecked())
		row = [self.file.text(), self.col.text(), self.col_name.text(), self.sheet_name.text(), self.Center.isChecked()]
		self.data.append(row)
		table = pandas.DataFrame(self.data, columns=["file", "col", "col_name", "sheet_name", "CENTER"])

		model = PandasModel(table)
		self.table.setModel(model)

	def setData(self, index, value):
		# print(self.data, value)
		self.data[index.row()][index.column()] = value

	def runOperation(self):
		print("RUN!")
		init_elec(Path = self.path.text() + "/", YearNum = int(self.YearNum.text()))
		# print(self.data)
		electricity(self.data)

# def main():
# 	app = QApplication([])
# 	window = MainWindow()
# 	window.show()

# 	app.exec()



if __name__ == '__main__':
    # main()
	app = QApplication([])
	window = MainWindow()
	window.show()

	app.exec()