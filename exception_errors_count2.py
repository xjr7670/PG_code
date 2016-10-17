#-*- coding:utf-8 -*-

#  本程序可以将指定目录下的所有htm文件进行统计
#  并将结果写到exception_errors.txt文件中
#  2016年10月13日

import os
import re
from bs4 import BeautifulSoup as bs


class Get_total(object):
    
    
    def __init__(self, file):
        self.f = open(file)
        self.html = self.f.read()
        exception_errors_start = self.html.find('Exception&nbsp;Errors')

        self.html = self.html[exception_errors_start:]
        self.html = self.html.replace("&nbsp;", " ")
        
        
    def get_total(self):
        pat_num = re.compile(r'\d{8}')
        s = pat_num.findall(self.html)
        total = len(s)
        return total


    def get_yellow(self):
        s = {
             "F00": 0,
             "F01": 0,
             "F02": 0,
             "F03": 0,
             "F04": 0,
             "F05": 0,
             "F06": 0,
             "F08": 0,
             "F09": 0
            }
        bsObj = bs(self.html, 'html.parser')
        nobr_tags =  bsObj.findAll('nobr', {'style': 'background:#f0f008'})
        for nobr in nobr_tags:
            text = nobr.get_text()
            res = text.split()

            if res[0] in s.keys():
                s[res[0]] += int(res[1])
            else:
                if len(res) == 2:
                    s[res[0]] = int(res[1])

        return s

    
    
if __name__ == "__main__":

    # working directory
    os.chdir('''G:\pg\SIPP Jul'16''')

    files = [x for x in os.listdir() if x.endswith('.htm')]

    f = open('exception_errors_count.txt', 'w')

    for eachFile in files:
        # the file need to count
        gt = Get_total(eachFile)
        
        total = gt.get_total()
        f.write("File of %s\n" % eachFile)
        f.write("The total of exception errors is: %s \n" % total)
        f.write("===========================================\n")

        y = gt.get_yellow()
        
        # print table head
        f.write("F00" + '\t' + "F01" + '\t' + "F02" + '\t' + "F03" + '\t' + "F04" + '\t' + "F05" + '\t' + "F06" + '\t' + "F08" + '\t' + "F09" + '\n')

        # formating output
        f.write(str(y['F00']) + '\t' + str(y['F01']) + '\t' + str(y['F02']) + '\t' + str(y['F03']) + \
              '\t' + str(y['F04']) + '\t' + str(y['F05']) + '\t' + str(y['F06']) + '\t' + str(y['F08']) + \
              '\t' + str(y['F09']) + '\n')
        f.write('\n\n\n\n')
    f.close()