#-*- coding:utf-8 -*-

#  本程序可以将指定目录下的所有htm文件进行统计
#  统计结果包含past due和exception errors
#  并将结果写到count_results.txt文件中
#  2016年10月14日


import os
import re
import json
from time import sleep
from bs4 import BeautifulSoup as bs

class Get_past_due_total(object):
    
    
    def __init__(self, hfile):
        self.f = open(hfile)
        self.html = self.f.read()
        self.pase_due_end = self.html.find('Exception&nbsp;Errors')
        self.html = self.html[:self.pase_due_end]
        self.html = self.html.replace("&nbsp;", " ")

        
    def get_total(self):
        pat_num = re.compile(r'\.\d{3}')
        s = pat_num.findall(self.html)
        total = len(s)
        return total
    
    def get_detail(self, tag, address):
        F_tag = tag
        address = address

        t = 0
        f01 = self.html.find(F_tag)
        if f01 == -1:
            return t

        f02 = self.html.find("F0", f01+1)

        c = self.html[f01:f02]

        t += c.count(address)

        while f01 != -1:
            f01 = self.html.find(F_tag, f02)
            f02 = self.html.find("F0", f01+1)
            c = self.html[f01:f02]
            t += c.count(address)
            
        return t


class Get_exception_errors_total(object):
    
    
    def __init__(self, hfile):
        self.f = open(hfile)
        self.html = self.f.read()
        self.exception_errors_start = self.html.find('Exception&nbsp;Errors')

        self.html = self.html[self.exception_errors_start:]
        self.html = self.html.replace("&nbsp;", " ")
        
    def get_total(self):
        pat_num = re.compile(r'\d{8}')
        s = pat_num.findall(self.html)
        total = len(s)
        return total

    def get_detail(self):
        s = {
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


class Work_dir(object):

    def __init__(self, work_dir):
        self.work_dir = work_dir
        if self.work_dir == '' or not os.path.exists(self.work_dir.replace('\\', '\\\\')):
            print 'No valid folder directory exists after last usage.'

    def set_work_dir(self):
        if os.path.exists(self.work_dir):
            new_dir = raw_input("Use last folder(just press Enter) or a new one(enter that):\t")
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

    config_file = open('config.json', 'r+')
    config_str = config_file.read()
    config = json.loads(config_str)
     
    # work direcotry
    work_dir = config["work_dir"]
    # MRP_C list
    mrp_c_list = config["mrp_controler"]
    # MRP_element list
    mrp_element_list = config["mrp_element"]

    # confirm work directory
    wd_obj = Work_dir(work_dir)
    new_dir = wd_obj.set_work_dir()
    if work_dir != new_dir:
        config['work_dir'] = new_dir
        config_file.write(config)
    os.chdir(work_dir)
    config_file.close()

    files = [x for x in os.listdir(work_dir) if x.endswith('.htm')]

    result_files = open('count_results.txt', 'w')

    for eachFile in files:

        # ************************************************************************************************************************
        # line 135-151 is use to counting exception errors information
        #
        # the file need to count
        exception_errors = Get_exception_errors_total(eachFile)
        
        total = exception_errors.get_total()
        result_files.write("Result of %s\n" % eachFile)
        result_files.write("The total of Exception errors is: %s \n" % total)
        result_files.write("===========================================\n")

        y = exception_errors.get_detail()
        
        # print table head
        result_files.write("F01" + '\t' + "F02" + '\t' + "F03" + '\t' + "F04" + '\t' + "F05" + '\t' + "F06" + '\t' + "F08" + '\t' + "F09" + '\n')

        # formating output
        result_files.write(str(y['F01']) + '\t' + str(y['F02']) + '\t' + str(y['F03']) + \
              '\t' + str(y['F04']) + '\t' + str(y['F05']) + '\t' + str(y['F06']) + '\t' + str(y['F08']) + \
              '\t' + str(y['F09']) + '\n')
        result_files.write('\n')
        result_files.flush()

        # ***************************************************************************************************************************
        # this code is use to counting past due information
        # the file need to count
        pase_due = Get_past_due_total(eachFile)
        
        total = pase_due.get_total()
      
        # f.write('File of ' + eachFile + '\n')
        result_files.write('The total of Past due is: ' + str(total) + '\n')
        result_files.write('=============================================================================================\n')
        
        d = {}
        
        # get all result number and put it to a dict variables
        for tag in mrp_c_list:
            d2 = {}
            for addr in mrp_element_list:
                t = pase_due.get_detail(tag, addr)
                d2[addr] = t
            d[tag] = d2
        
        # print table head
        result_files.write('\t\t' + "F01" + '\t' + "F02" + '\t' + "F03" + '\t' + "F04" + \
                '\t' + "F05" + '\t' + "F06" + '\t' + "F08" + '\t' + "F09" + '\t\n')
        
        # formating output
        for i in mrp_element_list:
            result_files.write(i + "\t" + str(d["F01"][i]) + "\t" + str(d["F02"][i]) + "\t" \
                      + str(d["F03"][i]) + "\t" + str(d["F04"][i]) + "\t" + str(d["F05"][i]) + "\t" + str(d["F06"][i]) + "\t" \
                      + str(d["F08"][i]) + "\t" + str(d["F09"][i]) + '\n')

        result_files.write('\n\n\n\n')
        result_files.flush()

    result_files.close()

    print '\nFinish!'
    print 'You can open the %s to see the result.' % os.path.abspath('count_results.txt')
    sleep(3)