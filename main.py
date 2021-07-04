# -*- coding:utf-8 -*-  
# __author__ = Damian
# __time__ = '2021/6/27 17:40'
# __project__ = 'ExcelTools'
from ui.call_main import MainWindow
from PyQt5 import QtGui, QtWidgets
import sys

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon(":/icon/xlsx.ico"))
    ExcelConverter = MainWindow()
    ExcelConverter.show()
    sys.exit(app.exec_())
