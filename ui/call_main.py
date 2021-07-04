# -*- coding:utf-8 -*-  
# __author__ = Damian
# __time__ = '2021/6/27 18:43'
# __project__ = 'ExcelTools'
import os
import inspect
from loguru import logger
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal, QMutex
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QWidget, QCheckBox, QListWidgetItem, QCompleter
from ui.ui_output.main import Ui_MainWindow
from PyQt5.QtCore import pyqtSignal, Qt
import ui.static.icons_rc
import threading
import time
from util_tools import cut_dict, array_split, get_file
import pythoncom
import win32com.client as win32


class ConverterForPyQT(QThread):
    logs_file_name = pyqtSignal(str)
    logs_file_count = pyqtSignal(int)

    def __init__(self, app_id_list, file_dict_list, mutex):
        super(ConverterForPyQT, self).__init__()
        self.app_id_list = app_id_list
        self.file_dict_list = file_dict_list
        self._mutex = mutex
        self.thread_pool = []
        self.counter = 0
        self.total = sum([len(item) for item in self.file_dict_list])

    def convert(self, app_id, file_dict):
        pythoncom.CoInitialize()
        app = win32.Dispatch(
            pythoncom.CoGetInterfaceAndReleaseStream(app_id, pythoncom.IID_IDispatch)
        )
        for file_name, file_path in file_dict.items():
            # logger.debug(file_dict)
            time.sleep(1)
            new_file = os.path.join(os.path.split(file_path)[0], file_name + 'x')
            # logger.debug(new_file)
            wb = app.Workbooks.Open(file_path)
            wb.SaveAs(new_file, FileFormat=51)  # FileFormat = 51 is for .xlsx extension
            wb.Close()  # FileFormat = 56 is for .xls extension
            os.remove(file_path)
            # logger.success(f"{file_name + 'x'} success")
            # queue.put(file_name)
            self._mutex.lock()
            self.counter += 1
            self.logs_file_name.emit(f"{file_name}转换成功, {self.counter}/{self.total}")
            self.logs_file_count.emit(round((self.counter / self.total) * 100))
            self._mutex.unlock()
        time.sleep(1)
        app.Application.Quit()
        pythoncom.CoInitialize()

    def run(self):
        for i in range(len(self.file_dict_list)):
            self.thread_pool.append(
                threading.Thread(
                    target=self.convert,
                    kwargs={"app_id": self.app_id_list[i],
                            "file_dict": self.file_dict_list[i]
                            }
                )
            )
            logger.info(f"put {self.app_id_list[i]} into thread >> {self.file_dict_list[i]}")
        for thread in self.thread_pool:
            thread.start()
            time.sleep(1)
        for thread in self.thread_pool:
            thread.join()


# 这个类才是实际进行操作的类，所以要继承重写Converter的内容

class LogsThread(QThread):
    logs_file_name = pyqtSignal(str)
    logs_file_count = pyqtSignal(int)

    def __init__(self, file_list: [], mutex):
        super(LogsThread, self).__init__()
        self._mutex = mutex
        self.file_list = file_list
        self.thread_pool = []
        self.counter = 0

    def logs(self, file_list):
        # counter = 0
        for file in file_list:
            time.sleep(0.5)
            self._mutex.lock()

            self.logs_file_name.emit(f"{file}转换成功, {self.counter}/{len(self.file_list)}")
            self.logs_file_count.emit(self.counter)
            logger.info(f"{threading.currentThread()}, {file}, {self.counter}")
            self.counter += 1
            self._mutex.unlock()

    def run(self):
        self.thread_pool.append(
            threading.Thread(
                target=self.logs,
                kwargs={"file_list": [i for i in range(100) if not i % 2]}
            )

        )
        self.thread_pool.append(
            threading.Thread(
                target=self.logs,
                kwargs={"file_list": [i for i in range(100) if i % 2]}
            )
        )
        for t in self.thread_pool:
            t.start()
            time.sleep(1)
        for t in self.thread_pool:
            t.join()


class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)

        # 配置path
        current_path = inspect.getfile(inspect.currentframe())
        dir_name = os.path.dirname(current_path)
        file_abs_path = os.path.abspath(dir_name)
        self.list_path = os.path.split(file_abs_path)

        # 配置qss
        qss_path = os.path.join(self.list_path[0], r"ui\static\Ubuntu.qss")
        self.setStyleSheet(open(qss_path, "r").read())

        # 配置attrs
        self.excel_dir_path = ""
        self._mutex = QMutex()
        self.file_dict_list = []
        self.thread_num = 2
        self.app_id_list = []

        # 绑定槽函数
        self.open_src_dir.clicked.connect(self.onOpenDirClicked)
        self.start.clicked.connect(self.onStarting)
        self.open_output_dir.clicked.connect(self.onOpenOutput)

        # 进度条设置
        self.progressBar.setMinimum(0)
        self.progressBar.setMaximum(100)
        self.progressBar.setValue(0)

    def init_converter(self):
        # 判断路径是否存在
        if os.path.isdir(self.src_dir.text()):
            file_dict = get_file(self.src_dir.text())
        else:
            logger.warning("文件夹路径不存在")
            self.logs_text.setText(f"文件夹路径不存在")
            return

        # 根据文件数量决定线程数
        if len(file_dict) == 1:
            self.file_dict_list.append(file_dict)
            self.thread_num = 1
        elif len(file_dict) == 0:
            logger.warning(f"文件夹下没有xls文件")
            self.logs_text.setText(f"该文件夹下没有xls文件需要转换")
            return
        else:
            self.thread_num = 2
            self.file_dict_list = cut_dict(file_dict, self.thread_num)

        # 创建win32com app
        for _ in range(self.thread_num):
            app = win32.Dispatch('Excel.Application')
            app_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, app)
            self.app_id_list.append(app_id)
            logger.info(f"app_id: {app_id}")
            time.sleep(1)

        self.converter = ConverterForPyQT(self.app_id_list, self.file_dict_list, self._mutex)

    def onOpenDirClicked(self):
        open_src_dir = QFileDialog.getExistingDirectory(self, "选取函证Excel文件夹", self.list_path[0])
        self.src_dir.setText(open_src_dir)

    def onOpenOutput(self):
        start_directory = self.src_dir.text()
        os.startfile(start_directory)

    def onStarting(self):
        self.progressBar.setValue(0)
        self.init_converter()
        self.converter.start()
        self.converter.logs_file_name.connect(self.onDisplayLogs)
        self.converter.logs_file_count.connect(self.onDisplayProcess)

    def onDisplayLogs(self, logs_file_name):
        self.logs_text.setText(logs_file_name)

    def onDisplayProcess(self, logs_file_count):
        self.progressBar.setValue(logs_file_count)


if __name__ == '__main__':
    import sys

    current_path = inspect.getfile(inspect.currentframe())
    dir_name = os.path.dirname(current_path)
    file_abs_path = os.path.abspath(dir_name)
    list_path = os.path.split(file_abs_path)

    qss_path = os.path.join(list_path[0], r"ui\static\Ubuntu.qss")
    app = QtWidgets.QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon(":/icon/xlsx.png"))
    AutoConfirmation = MainWindow()
    AutoConfirmation.setStyleSheet(open(qss_path, "r").read())
    AutoConfirmation.show()
    sys.exit(app.exec_())
