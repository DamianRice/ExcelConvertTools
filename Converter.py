# -*- coding:utf-8 -*-  
# __author__ = Damian
# __time__ = '2021/6/25 21:25'
# __project__ = 'ExcelTools'
from loguru import logger
import win32com.client as win32
import os
import pythoncom
from tqdm import tqdm
import threading
import time
from queue import Queue
from util_tools import cut_dict, array_split, get_file


class Converter:
    def __init__(self, thread_num: int = 2):
        # 基本field
        self.thread_num = thread_num
        self.app = []
        self.app_id = []
        self.file_dict = {}
        self.file_dict_list = []

        # 初始化线程相关
        self.thread_pool = []
        self.queue = Queue(maxsize=0)

        # 初始化进程实例以及ID
        for _ in range(self.thread_num):
            app = win32.Dispatch('Excel.Application')
            app_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, app)
            self.app.append(app)
            self.app_id.append(app_id)
            logger.info(f"app_id: {self.app_id}")
            time.sleep(1)

    @staticmethod
    def get_file(self, path, suffix=".xls"):
        file_names = []
        file_paths = []
        for root, dirs, files in os.walk(path, topdown=False):
            for name in files:
                if not os.path.splitext(name)[0].startswith("~$"):
                    if os.path.splitext(name)[1] == suffix:
                        logger.debug(f"{name} was found")
                        file_names.append(name)
                        file_paths.append(os.path.join(root, name))

        file_dict = dict(zip(file_names, file_paths))
        # res = list(file_dict.items())
        # self.file_dict = file_dict
        logger.info(f"file_dict: {file_dict}")
        return file_dict

    def cut_dict(self, file_dict):

        file_dict_list = []
        _ = {}
        file_name = list(file_dict.keys())
        file_path = list(file_dict.values())

        cut_name = self.array_split(file_name, self.thread_num)
        cut_path = self.array_split(file_path, self.thread_num)

        for i in range(len(cut_name)):
            _ = dict(zip(cut_name[i], cut_path[i]))
            file_dict_list.append(_.copy())
            _.clear()
        self.file_dict_list = file_dict_list
        logger.info(f"file_dict_list: {self.file_dict_list}")
        return self.file_dict_list

    def dispatch(self, app_id_list, file_dict_list):
        for i in range(len(file_dict_list)):
            self.thread_pool.append(
                threading.Thread(
                    target=self.convert,
                    kwargs={"app_id": app_id_list[i],
                            "file_dict": file_dict_list[i],
                            "queue": self.queue,
                            }
                )
            )

            logger.info(f"put {app_id_list[i]} into thread >> {file_dict_list[i]}")

    def convert(self, app_id, file_dict, queue):
        pythoncom.CoInitialize()
        app = win32.Dispatch(
            pythoncom.CoGetInterfaceAndReleaseStream(app_id, pythoncom.IID_IDispatch)
        )
        for file_name, file_path in file_dict.items():
            time.sleep(1)
            new_file = os.path.join(os.path.split(file_path)[0], file_name + 'x')
            wb = app.Workbooks.Open(file_path)
            wb.SaveAs(new_file, FileFormat=51)  # FileFormat = 51 is for .xlsx extension
            wb.Close()  # FileFormat = 56 is for .xls extension
            os.remove(file_path)
            # logger.success(f"{file_name + 'x'} success")
            queue.put(file_name)
        time.sleep(1)
        app.Application.Quit()
        pythoncom.CoInitialize()

    def result(self, queue: Queue):
        while True:
            if queue.qsize() != 0:
                res = queue.get()
                logger.success(f"{res + 'x'} success")
            else:
                continue

    def run(self, path, suffix=".xls"):
        self.file_dict = self.get_file(path, suffix)

        # 只有一个文件
        if len(self.file_dict) == 1:
            self.file_dict_list.append(self.file_dict)
            self.thread_num = 1
        #
        logs_thread = threading.Thread(target=self.result, kwargs={"queue": self.queue, })
        self.thread_pool.append(logs_thread)
        self.file_dict_list = self.cut_dict(self.file_dict)
        self.dispatch(self.app_id, self.file_dict_list)
        for thread in self.thread_pool:
            thread.start()
            time.sleep(1)
        for thread in self.thread_pool:
            thread.join()
        # print(threading.active_count())
        # 附加的result进程可以根据如下三个进程进行判断结束
        print(threading.activeCount())
        print(threading.currentThread())
        print(threading.enumerate())

    @staticmethod
    def array_split(ary, indices_or_sections):
        if len(ary) < indices_or_sections:
            raise AttributeError
        Ntotal = len(ary)
        # indices_or_sections is a scalar, not an array.
        Nsections = int(indices_or_sections)
        if Nsections <= 0:
            raise ValueError('number sections must be larger than 0.')
        Neach_section, extras = divmod(Ntotal, Nsections)
        section_sizes = ([0] +
                         extras * [Neach_section + 1] +
                         (Nsections - extras) * [Neach_section])
        div_points = []
        for i in range(len(section_sizes)):
            if i == 0:
                div_points.append(section_sizes[0])
            else:
                div_points.append(sum(section_sizes[:i + 1]))

        sub_arys = []
        for i in range(Nsections):
            st = div_points[i]
            end = div_points[i + 1]
            sub_arys.append(ary[st:end])

        return sub_arys


class ConverterForQT:
    def __init__(self):
        pass


if __name__ == '__main__':
    converter = Converter(thread_num=2)
    converter.run(path=r"C:\Users\mi007\Desktop\AutoPwC\ExcelTools\temp")
