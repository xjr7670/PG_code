# -*- coding:utf-8 -*-

#  本程序统计指定htm文件，是原程序的单文件版本
#  在使用时，需要指定文件路径
#  统计结果包含past due和exception errors
#  统计结果直接输出
#  2016年12月30日


import os
import re
from tkinter import filedialog
import sys
from time import sleep
from bs4 import BeautifulSoup as bs


class Get_past_due_total(object):
    '''
    统计past due，每次接收一个文件进行处理
    截取从开头至Exception Error处的内容进行统计
    '''

    def __init__(self, error_file):
        self.f = open(error_file)
        self.html = self.f.read()
        self.pase_due_end = self.html.find('Exception&nbsp;Errors')
        self.html = self.html[:self.pase_due_end]
        self.html = self.html.replace("&nbsp;", " ")

    def get_total(self):
        '''
        利用统计错误数量的方式统计总数
        数量是以点号加3个数字结束的
        把统计得到的总数返回，返回变量的类型为int
        '''

        pat_num = re.compile(r'\.\d{3}')
        s = pat_num.findall(self.html)
        total = len(s)
        return total

    def get_detail(self, mrp_c, mrp_e):
        '''
        函数接收一个mrp_c和一个mrp_e，每次在一个文件中，
        统计所接收的mrp_c和下一个mrp_c中间，有多少个mrp_e。
        每次只统计一个mrp_c的一个mrp_e结果总数，并返回，类型为int
        '''
        MPR_c = mrp_c
        MRP_e = mrp_e

        total = 0
        f01 = self.html.find(MPR_c)
        if f01 == -1:
            return total

        f02 = self.html.find("F0", f01 + 1)

        content = self.html[f01:f02]

        total += content.count(MRP_e)

        while f01 != -1:
            f01 = self.html.find(MPR_c, f02)
            f02 = self.html.find("F0", f01 + 1)
            content = self.html[f01:f02]
            total += content.count(MRP_e)

        return total


class Get_exception_errors_total(object):
    '''
    统计past due，每次接收一个文件进行处理
    截取从Exception Error至结尾的内容进行统计
    '''

    def __init__(self, error_file):
        self.f = open(error_file)
        self.html = self.f.read()
        self.exception_errors_start = self.html.find('Exception&nbsp;Errors')

        self.html = self.html[self.exception_errors_start:]
        self.html = self.html.replace("&nbsp;", " ")

    def get_total(self):
        '''
        利用统计metirial数量的方式统计总数
        metirial是以8个数字为代号表示的
        把统计得到的总数返回，类型为int
        '''

        pat_num = re.compile(r'\d{8}')
        s = pat_num.findall(self.html)
        total = len(s)
        return total

    def get_detail(self, mrp_c_list):
        '''
        利用DOM来获取含有mrp_c和出错数量的标签
        这个标签有个唯一的style为background:#f0f008
        用字典表示mrp_c和出错数量
        用空格分割其中的文本，把出错数量分别加到以mrp_c为键的值中
        返回统计结果，类型为dict
        '''

        result_dict = dict()
        for each_mrp_c in mrp_c_list:
            result_dict[each_mrp_c] = 0

        bsObj = bs(self.html, 'html.parser')
        nobr_tags = bsObj.findAll('nobr', {'style': 'background:#f0f008'})
        for nobr in nobr_tags:
            text = nobr.get_text()
            res = text.split()

            if res[0] in result_dict.keys():
                result_dict[res[0]] += int(res[1])
            else:
                if len(res) == 2:
                    result_dict[res[0]] = int(res[1])
        return result_dict



if __name__ == "__main__":

    # 打开配置文件，获取配置信息并转成dict
    # 配置信息包含路径、mrp_controller列表、mrp_element列表
    if not os.path.exists('config.cfg'):
        print("There is no configuration file *config.cfg* in current folder!")
        os._exit(0)

    config_file = open('config.cfg')
    config_str = config_file.read()
    config_str = re.sub(r'[\n\t]', '', config_str)
    config = eval(config_str)

    # 获取mrp_controller列表
    mrp_c_list = config["mrp_controller"]
    # 获取mrp_element列表
    mrp_element_list = config["mrp_element"]
    config_file.close()

    # 获取文件
    error_file = config["error_file"]

    # 判断文件是否存在
    if error_file == "" or not os.path.exists(error_file):
        # 使用win32ui模块获取文件
        error_file = filedialog.askopenfilename(defaultextension = "htm", initialdir = "C:/")
        config["error_file"] = error_file
        # 把文件路径再写入
        with open('config.cfg', 'w') as f:
            f.write(str(config))
    

    if error_file == "":
        print("You haven't choose any file")
        sleep(3)
        sys.exit()

    result_file = open("result.txt", "w")
    result_file.write("Result of %s" % error_file + "\n\n")
    # ***************************************************************************************************************************
    # this code is use to counting past due information

    # the file need to count
    pase_due = Get_past_due_total(error_file)

    total = pase_due.get_total()


    # f.write('File of ' + eachFile + '\n')
    result_file.write('\nThe total of Past due is: ' + str(total))
    result_file.write('\n=============================================================================================\n')

    d = {}

    # get all result number and put it to a dict variables
    for each_mrp_c in mrp_c_list:
        d2 = {}
        for each_mrp_e in mrp_element_list:
            each_total = pase_due.get_detail(each_mrp_c, each_mrp_e)
            d2[each_mrp_e] = each_total
        d[each_mrp_c] = d2

    # print table head
    past_due_table_head = '\t\t'
    for each_mrp_c in mrp_c_list:
        past_due_table_head += each_mrp_c + '\t'
    past_due_table_head += '\n'
    result_file.write(past_due_table_head + "\n")

    # formating output
    for each_element in mrp_element_list:
        res = each_element
        for each_mrp_c in mrp_c_list:
            res += "\t" + str(d[each_mrp_c][each_element])
        result_file.write(res + "\n")



    # ************************************************************************************************************************
    # this code is use to counting exception errors information

    # the file need to count
    exception_errors = Get_exception_errors_total(error_file)

    total = exception_errors.get_total()
    result_file.write("\nThe total of Exception errors is: %s \n" % total)
    result_file.write("===========================================\n")

    exception_error_result_dict = exception_errors.get_detail(mrp_c_list)

    # print table head
    exception_errors_table_head = ''
    for each_mrp_c in mrp_c_list:
        exception_errors_table_head += each_mrp_c + '\t'
    result_file.write(exception_errors_table_head + "\n")

    # formating output
    exception_errors_string = ''
    for each_mrp_c in mrp_c_list:
        exception_errors_string += str(exception_error_result_dict[each_mrp_c]) + '\t'
    result_file.write(exception_errors_string + "\n")
    result_file.close()
    print("Please open result.txt to get the result")
    sleep(5)