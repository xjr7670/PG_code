# -*- coding:utf-8 -*-

#  本程序可以将指定目录下的所有htm文件进行统计
#  统计结果包含past due和exception errors
#  并将结果写到count_results.txt文件中
#  2016年10月14日


import os
import re
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


class Work_dir(object):
    '''
    接收配置文件中的工作路径，并检查是否有效
    如果为空或路径无效，则要求输入文件夹路径
    否则询问是使用前一次的路径还是重新输入
    所有输入的路径都会经过一次路径有效性的检查
    '''

    def __init__(self, work_dir):
        self.work_dir = work_dir
        if self.work_dir == '' or not os.path.exists(self.work_dir):
            print 'No valid folder directory exists after last usage.'

    def set_work_dir(self):
        '''
        检查文件夹路径是有效，如果有效则询问是否使用新路径
        如果使用新路径，则返回新路径，类型为string
        否则返回所传入的配置文件中的路径，类型为string
        '''

        if os.path.exists(self.work_dir):
            new_dir = raw_input("Use last folder %s (just press Enter) or a new one(enter that):\t" % self.work_dir)
            if new_dir != '':
                new_dir = self.confirm_dir(new_dir)
                return new_dir
            else:
                return self.work_dir
        else:
            new_dir = raw_input("Please enter the working directory: ")
            new_dir = self.confirm_dir(new_dir)
            return new_dir

    def confirm_dir(self, work_dir):
        '''
        接收一个路径字符串，检查路径是否有效
        有效，则直接返回
        无效，则要求重新输入，并把经过检测后的新路径返回
        '''

        if os.path.exists(work_dir):
            return work_dir
        else:
            while not os.path.exists(work_dir):
                work_dir = raw_input("Please enter the correct folder path, start with the C/D/E etc., "
                                     "Case insensitive: ")
                work_dir.strip()
            else:
                return work_dir


if __name__ == "__main__":

    # 打开配置文件，获取配置信息并转成dict
    # 配置信息包含路径、mrp_controller列表、mrp_element列表
    config_file = open('config.cfg')
    config_str = config_file.read()
    config_str = re.sub(r'[\n\t]', '', config_str)
    config = eval(config_str)

    # 获取工作目录路径
    work_dir = config["work_dir"]
    # 获取mrp_controller列表
    mrp_c_list = config["mrp_controller"]
    # 获取mrp_element列表
    mrp_element_list = config["mrp_element"]
    config_file.close()

    # confirm work directory
    wd_obj = Work_dir(work_dir)
    new_dir = wd_obj.set_work_dir()
    if work_dir != new_dir:
        config['work_dir'] = new_dir
        with open('config.cfg', 'w') as config_file:
            config_file.write(str(config))
            work_dir = new_dir
    os.chdir(work_dir)

    # 获取工作目录下的所有htm文件名并放到列表中
    files = [x for x in os.listdir(work_dir) if x.endswith('.htm')]

    result_file = open('count_results.txt', 'w')

    for eachFile in files:

        result_file.write("Result of %s\n" % eachFile)
        # ***************************************************************************************************************************
        # this code is use to counting past due information
        
        # the file need to count
        pase_due = Get_past_due_total(eachFile)

        total = pase_due.get_total()

        # f.write('File of ' + eachFile + '\n')
        result_file.write('The total of Past due is: ' + str(total) + '\n')
        result_file.write('=============================================================================================\n')

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
        result_file.write(past_due_table_head)

        # formating output
        for each_element in mrp_element_list:
            res = each_element
            for each_mrp_c in mrp_c_list:
                res += "\t" + str(d[each_mrp_c][each_element])
            res += '\n'
            result_file.write(res)

        result_file.write('\n')
        result_file.flush()

        # ************************************************************************************************************************
        # this code is use to counting exception errors information

        # the file need to count
        exception_errors = Get_exception_errors_total(eachFile)

        total = exception_errors.get_total()
        result_file.write("The total of Exception errors is: %s \n" % total)
        result_file.write("===========================================\n")

        exception_error_result_dict = exception_errors.get_detail(mrp_c_list)

        # print table head
        exception_errors_table_head = ''
        for each_mrp_c in mrp_c_list:
            exception_errors_table_head += each_mrp_c + '\t'
        exception_errors_table_head += '\n'
        result_file.write(exception_errors_table_head)

        # formating output
        exception_errors_string = ''
        for each_mrp_c in mrp_c_list:
            exception_errors_string += str(exception_error_result_dict[each_mrp_c]) + '\t'
        exception_errors_string += '\n'
        result_file.write(exception_errors_string)

        result_file.flush()
        result_file.write('\n\n\n\n')

    result_file.close()

    print '\nFinish!'
    print 'You can open the %s to see the result.' % os.path.abspath('count_results.txt')
    sleep(3)