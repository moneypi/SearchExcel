import sys, os
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QAbstractItemView, QTableWidgetItem, QTableWidget, \
    QFileDialog
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot, QCoreApplication, Qt
from PyQt5.Qt import QLineEdit
import xlrd

_translate = QCoreApplication.translate


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'SearchExcel '

        self.pathTextBox = QLineEdit(self)
        self.setWindowTitle(self.title)

        self.pathTextBox.move(20, 20)

        # create searchTextBox
        self.searchTextBox = QLineEdit(self)
        self.searchTextBox.move(20, 80)
        self.searchTextBox.textChanged.connect(self.onTextChange)

        # Create a openPathButton in the window
        self.openPathButton = QPushButton('打开文件夹', self)

        self.tableWidget = QTableWidget(self)

        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(5)
        self.tableWidget.setRowCount(0)

        for i in range(3):
            item = QTableWidgetItem()
            self.tableWidget.setVerticalHeaderItem(i, item)

        for i in range(5):
            item = QTableWidgetItem()
            item.setTextAlignment(Qt.AlignCenter)
            self.tableWidget.setHorizontalHeaderItem(i, item)
        self.tableWidget.horizontalHeader().setCascadingSectionResizes(True)

        item = self.tableWidget.horizontalHeaderItem(0)
        item.setText(_translate("widget", "文件名"))
        item = self.tableWidget.horizontalHeaderItem(1)
        item.setText(_translate("widget", "Sheet"))
        item = self.tableWidget.horizontalHeaderItem(2)
        item.setText(_translate("widget", "行"))
        item = self.tableWidget.horizontalHeaderItem(3)
        item.setText(_translate("widget", "列"))
        item = self.tableWidget.horizontalHeaderItem(4)
        item.setText(_translate("widget", "内容"))
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # connect openPathButton to function on_click
        self.openPathButton.clicked.connect(self.on_click)
        self.showMaximized()

    @pyqtSlot()
    def on_click(self):
        self.pathTextBox.setText(QFileDialog.getExistingDirectory(None, "请选择文件夹路径"))

    def resizeEvent(self, a0):
        self.pathTextBox.resize(self.width() * 0.8, 40)
        self.openPathButton.move(self.width() * 0.85, 20)
        self.openPathButton.resize(self.width() * 0.1, 40)
        self.searchTextBox.resize(self.width() - 40, 40)
        self.tableWidget.setGeometry(20, 180, self.width() - 40, self.height() - 200)
        self.tableWidget.horizontalHeader().resizeSection(0, int(self.tableWidget.width() * 0.3))
        self.tableWidget.horizontalHeader().resizeSection(1, int(self.tableWidget.width() * 0.2))
        self.tableWidget.horizontalHeader().resizeSection(2, int(self.tableWidget.width() * 0.05))
        self.tableWidget.horizontalHeader().resizeSection(3, int(self.tableWidget.width() * 0.05))
        self.tableWidget.horizontalHeader().resizeSection(4, int(self.tableWidget.width() * 0.4))

    def onTextChange(self):
        result_list = []
        searchText = self.searchTextBox.text()
        if len(searchText) == 0:
            self.tableWidget.setRowCount(0)
            return
        excel_dir_path = self.pathTextBox.text()
        file_list = os.listdir(excel_dir_path)
        if file_list is None:
            self.tableWidget.setRowCount(0)
            return
        for file_name in file_list:
            if file_name.endswith("xlsx") or file_name.endswith("xls"):
                excel_file_path = os.path.join(excel_dir_path, file_name)
                try:
                    excel = xlrd.open_workbook(excel_file_path, encoding_override="utf-8")
                except IOError:
                    print("open %s failed" % excel_file_path)
                else:
                    all_sheet = excel.sheet_names()
                    for sheet_name in all_sheet:
                        each_sheet_by_name = excel.sheet_by_name(sheet_name)
                        for i in range(each_sheet_by_name.nrows):
                            for j in range(each_sheet_by_name.ncols):
                                if searchText in str(each_sheet_by_name.row_values(i)[j]):
                                    result_list.append(
                                        (file_name, sheet_name, i + 1, j + 1, each_sheet_by_name.row_values(i)[j]))

        self.tableWidget.setRowCount(len(result_list))

        for i in range(len(result_list)):
            for j in range(5):
                item = QTableWidgetItem()
                self.tableWidget.setItem(i, j, item)
                item.setText(_translate("widget", str(result_list[i][j])))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    app.exit(app.exec_())
