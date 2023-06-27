# -*- coding: UTF-8 -*-
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import datetime
import csv


class bug_analyse():
    def __init__(self):
        self.total_bug_count = 0

    def read_csv_as_dict(self, file_paths=[], encoding='utf-8'):
        result = []
        for file_path in file_paths:
            with open(file_path, 'r', encoding=encoding) as csv_file:
                reader = csv.reader(csv_file)
                headers = next(reader)  # 读取首行作为键
                # for row in reader:
                #     values = row
                #     entry = dict(zip(headers, values))
                #     result[row[0]] = entry
                for row in reader:
                    entry = dict(zip(headers, row))
                    # zip(headers, row) 创建了一个迭代器，它将 headers 和 row 中的对应元素一一配对。
                    # 然后，dict() 函数将配对的元素转换为字典的键值对，其中 headers 中的元素作为键，
                    # row 中的元素作为相应的值。这样就生成了一个字典，其中包含了一行数据的键值对。
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
            print(converted_date.strftime('%Y-%m-%d'))
            # print(type(converted_date))
            return converted_date
        except ValueError:
            return None

    def get_bug_by_date(self, startdate: datetime.datetime, enddate: datetime.datetime, buglist=[]):
        """通过日期区间筛选bug列表,默认参数必须放在非默认参数之后"""
        need_bug_list = []
        for bug in buglist:
            bug_create_date = bug["创建日期"]
            bug_create_date = bug_create_date.split()[0]
            bug_create_date = datetime.datetime.strptime(bug_create_date, '%Y-%m-%d')
            if startdate <= bug_create_date <= enddate:
                need_bug_list.append(bug)
        # print(need_bug_list)
        return need_bug_list

    def bug_level_analyse(self, buglist=[]):
        """bug等级分析"""
        leveldict = {}
        for bug in buglist:
            level = bug['严重程度']
            if level not in leveldict:
                leveldict[level] = 1
            else:
                leveldict[level] += 1
        print("BUG按严重程度分类，其中:")
        for level in leveldict:
            print("{}{}个，占比{}；".format(level, leveldict[level], "{:.1f}%".format(
                leveldict[level] / self.total_bug_count * 100)))
        print("************************************************************")

    def bug_online_analyse(self, buglist=[]):
        """bug线上线下分析"""
        online = 0
        offline = 0
        for bug in buglist:
            level = bug['Bug标题']
            if '线上-' in level or '正式-' in level:
                online += 1
            else:
                offline += 1
        print("BUG线上线下分类，其中线上{}个，占比{}；线下{}个，占比{}".format(online, "{:.1f}%".format(
            online / self.total_bug_count * 100), offline, "{:.1f}%".format(offline / self.total_bug_count * 100)))
        print("************************************************************")

    def bug_resolution_analyse(self, buglist=[]):
        """bug解决方案分析"""
        resolution = {}
        resolve_num = 0
        for bug in buglist:
            bugstatus = bug['Bug状态']
            if '关闭' in bugstatus or '解决' in bugstatus:
                res = bug['解决方案']
                if res not in resolution:
                    resolution[res] = 1
                else:
                    resolution[res] += 1
                resolve_num += 1
        print("BUG按解决方案分类，其中:")
        for res in resolution:
            print("{}{}个，占比{}；".format(res, resolution[res], "{:.1f}%".format(
                resolution[res] / resolve_num * 100)))
        print("************************************************************")

    def bug_responsible_analyse(self, buglist=[]):
        """bug责任人分析"""
        responsible = {}
        for bug in buglist:
            bugstatus = bug['Bug状态']
            if '关闭' in bugstatus or '解决' in bugstatus:
                res = bug['解决者']
            else:
                res = bug['指派给']
            if res not in responsible:
                responsible[res] = 1
            else:
                responsible[res] += 1
        print("BUG按责任人统计，其中:")
        for res in responsible:
            print("{}{}个，占比{}；".format(res, responsible[res], "{:.1f}%".format(
                responsible[res] / self.total_bug_count * 100)))
        print("************************************************************")

    def bug_reopen_analyse(self, buglist=[]):
        """bug反复激活分析"""
        print("BUG多次激活列表:")
        for bug in buglist:
            bugreopentime = bug['激活次数']
            if bugreopentime == 0 or bugreopentime == '0':
                pass
            else:
                bugid = bug['Bug编号']
                bugtitle = bug['Bug标题']
                buglevel = bug['严重程度']
                bugstatus = bug['Bug状态']
                if '关闭' in bugstatus or '解决' in bugstatus:
                    bugresponsible = bug['解决者']
                else:
                    bugresponsible = bug['指派给']
                print("{} {} 责任人：{} 严重程度：{} 激活次数：{}；".format(bugid, bugtitle, bugresponsible, buglevel, bugreopentime))
        print("************************************************************")

    def total_analyse(self):
        filepath = self.select_file()
        buglist = self.read_csv_as_dict(filepath)
        startdate = self.convert_date_format("请输入开始日期，格式可以是月/日或月-日或月.日：")
        enddate = self.convert_date_format("请输入结束日期，格式可以是月/日或月-日或月.日：")
        buglist = self.get_bug_by_date(startdate, enddate, buglist)
        self.total_bug_count = len(buglist)
        print("{}至{}区间内共发现{}个bug".format(startdate.strftime('%Y-%m-%d'), enddate.strftime('%Y-%m-%d'),
                                                 self.total_bug_count))
        self.bug_level_analyse(buglist)
        self.bug_online_analyse(buglist)
        self.bug_resolution_analyse(buglist)
        self.bug_responsible_analyse(buglist)
        self.bug_reopen_analyse(buglist)


if __name__ == "__main__":
    total_analyse = bug_analyse()
    total_analyse.total_analyse()
