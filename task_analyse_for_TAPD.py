# -*- coding: UTF-8 -*-
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import datetime
# import csv
import openpyxl
import jieba
from operator import itemgetter


class task_analyse():
    """使用前先将TAPD中导出的xls文件，另存为xlsx文件
    用来分析TAPD中任务工时统计，需要2张表
    1.按时间区间筛选出来的已完成状态的需求（故事）表
    2.已完成状态的任务表"""
    def __init__(self):
        self.task_on_person = {}  #用户存储个人维度的bug数据，内部结构{'name': hours, }

    def read_xls_as_dict(self, file_paths=[]):
        result = []
        for file_path in file_paths:
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active

            # 读取首行作为键
            headers = [cell.value for cell in sheet[1]]

            # 逐行读取数据并转换为字典
            for row in sheet.iter_rows(min_row=2, values_only=True):
                entry = dict(zip(headers, row))
                result.append(entry)
        return result

    def select_file(self):
        filepath = filedialog.askopenfilenames()
        return filepath

    def convert_date_format(self, notice="string"):
        """将输入时间转化成标准格式时间"""
        date_string = input(notice)
        current_year = datetime.datetime.now().year
        date_parts = date_string.split('/')
        if len(date_parts) != 2:
            date_parts = date_string.split('-')
        if len(date_parts) != 2:
            date_parts = date_string.split('.')
        if len(date_parts) != 2:
            date_parts = None
        month = int(date_parts[0])
        day = int(date_parts[1])
        try:
            converted_date = datetime.datetime(current_year, month, day)
            # print(converted_date.strftime('%Y-%m-%d'))
            # print(type(converted_date))
            return converted_date
        except ValueError:
            return None

    def get_storys_id_list(self, storys_list=[]):
        """获取需要统计的故事id清单"""
        id_list=[]
        for story in storys_list:
            storyid = story["ID"]
            storyid = storyid.split()[0]
            id_list.append(storyid)
        # print(id_list)
        return id_list

    def analyse_task(self, task_list=[], story_list=[], only_this_month_story=True):
        """任务分析，统计到个人"""
        for task in task_list:
            story_id = task["需求ID"]
            estimated_hour = task["预估工时"]
            if estimated_hour == '--':
                estimated_hour = 0
            elif estimated_hour == None:
                estimated_hour = 0
            estimated_hour = float(estimated_hour)
            transactor = task["处理人"]
            task_desc = task["标题"]
            task_desc = task_desc + " 工时：" + str(estimated_hour)
            if only_this_month_story:
                if story_id in story_list:
                    if transactor not in self.task_on_person:
                        self.task_on_person[transactor] = {'工时': estimated_hour, '任务': [task_desc]}
                        # self.task_on_person[transactor]['工时'] = estimated_hour
                        # self.task_on_person[transactor]['任务'] = []
                        # self.task_on_person[transactor]['任务'].append(task_desc)
                    else:
                        self.task_on_person[transactor]['工时'] += estimated_hour
                        self.task_on_person[transactor]['任务'].append(task_desc)
            else:
                if transactor not in self.task_on_person:
                    self.task_on_person[transactor] = {'工时': estimated_hour, '任务': [task_desc]}
                    # self.task_on_person[transactor]['工时'] = estimated_hour
                    # self.task_on_person[transactor]['任务'] = []
                    # self.task_on_person[transactor]['任务'].append(task_desc)
                else:
                    self.task_on_person[transactor]['工时'] += estimated_hour
                    self.task_on_person[transactor]['任务'].append(task_desc)
        return self.task_on_person


    def total_analyse(self, need_story = False):
        """先读取完成的故事列表的excel，再读取任务的excel"""
        if need_story:
            filepath = self.select_file()
            storylist = self.read_xls_as_dict(filepath)
            storyidlist = self.get_storys_id_list(storylist)
        taskfilepath = self.select_file()
        tasklist = self.read_xls_as_dict(taskfilepath)
        if need_story:
            self.analyse_task(tasklist, storyidlist, False)
        else:
            self.analyse_task(tasklist, [], False)
        for user in self.task_on_person:
            if user not in ('XXX'):
                print("{}工时：{}".format(user, self.task_on_person[user]['工时']))
        for user in self.task_on_person:
            if user not in ('XXX'):
                print("{}工时：{}".format(user, self.task_on_person[user]['工时']))
                for task in self.task_on_person[user]['任务']:
                    print(task)
                    


if __name__ == "__main__":
    total_analyse = task_analyse()
    total_analyse.total_analyse()
