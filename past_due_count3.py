#-*- coding:utf-8 -*-

#  本程序可以将指定目录下的所有htm文件进行统计
#  并将结果写到past_due_count.txt文件中
#  2016年10月13日


import os
import re
import pprint

class Get_total(object):
    
    
    def __init__(self, file):
        self.f = open(file)
        self.html = self.f.read()
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
    
if __name__ == "__main__":

    # working directory
    os.chdir('''G:\pg\SIPP Jul'16''')

    f = open('past_due_count.txt', 'w')

    files = [x for x in os.listdir() if x.endswith('.htm')]

    # MRP_C list
    tag_list = ["F00", "F01", "F02", "F03", "F04", "F05", "F06", "F08", "F09"]

    # MRP_element list
    l2 = ['AR Order.res.',
          'BA PurRequist',
          'BE Purch.ord.',
          'BR Proc.order',
          'LA Ship.note',
          'LB StLoc.stck',
          'LE Del.sched.',
          'PA Plnd order',
          'QM QM InspLot',
          'SB Depend.req',
          'U1 Rel. order',
          'U2 PchReq.Rel',
          'U3 PldOrdRel.',
          'U4 SAgmt rel.',
          'VJ Delivery'
          ]


    for eachFile in files:
        # the file need to count
        gt = Get_total(eachFile)
        
        total = gt.get_total()
      
        f.write('File of ' + eachFile + '\n')
        f.write('The total number is: ' + str(total) + '\n')
        f.write('=============================================================================================\n')
        
        d = {}
        
        for tag in tag_list:
            d2 = {}
            for addr in l2:
                t = gt.get_detail(tag, addr)
                d2[addr] = t
            d[tag] = d2

        l1 = list(d.keys())
        l1.sort()
        l2 = list(d['F01'].keys())
        l2.sort()
        
        # print table head
        f.write('\t\t' + "F00" + '\t' + "F01" + '\t' + "F02" + '\t' + "F03" + '\t' + "F04" + \
                '\t' + "F05" + '\t' + "F06" + '\t' + "F08" + '\t' + "F09" + '\t\n')
        
        # formating output
        for i in l2:
            f.write(i + "\t" + str(d["F00"][i]) + "\t" + str(d["F01"][i]) + "\t" + str(d["F02"][i]) + "\t" \
                      + str(d["F03"][i]) + "\t" + str(d["F04"][i]) + "\t" + str(d["F05"][i]) + "\t" + str(d["F06"][i]) + "\t" 
                      + str(d["F08"][i]) + "\t" + str(d["F09"][i]) + '\n')

        f.write('\n\n\n\n')

    f.close()