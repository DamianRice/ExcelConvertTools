# -*- coding:utf-8 -*-  
# __author__ = Damian
# __time__ = '2021/7/4 13:14'
# __project__ = 'ExcelTools'
import os
from loguru import logger

__all__ = ["get_file", "array_split", "cut_dict"]


def get_file(path, suffix=".xls"):
    file_names = []
    file_paths = []
    for root, dirs, files in os.walk(path, topdown=False):
        for name in files:
            if not os.path.splitext(name)[0].startswith("~$"):
                if os.path.splitext(name)[1] == suffix:
                    logger.debug(f"{name} was found")
                    file_names.append(name)
                    file_paths.append(os.path.abspath(os.path.join(root, name)))

    file_dict = dict(zip(file_names, file_paths))
    # res = list(file_dict.items())
    # self.file_dict = file_dict
    logger.info(f"file_dict: {file_dict}")
    return file_dict


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


def cut_dict(file_dict, num):
    file_dict_list = []
    _ = {}
    file_name = list(file_dict.keys())
    file_path = list(file_dict.values())

    cut_name = array_split(file_name, num)
    cut_path = array_split(file_path, num)

    for i in range(len(cut_name)):
        _ = dict(zip(cut_name[i], cut_path[i]))
        file_dict_list.append(_.copy())
        _.clear()
    file_dict_list = file_dict_list
    logger.info(f"file_dict_list: {file_dict_list}")
    return file_dict_list
