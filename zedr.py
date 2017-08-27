# -*- coding: utf-8 -*-
"""
Created on Sun Aug 27 14:35:27 2017

@author: Administrator
"""

from lxml import etree

def get_data(fname, code_list):
    """接受文件名和Code列表来处理数据"""
    
    src = open(fname).read()
    htm = etree.HTML(src)
    p = htm.xpath("//blockquote/p[2]")[0]
    trs = p.findall(".//tr")
    # 开头行是表头，结尾行是表尾，中间还有位于第2页的表头，是第41个tr
    # 这3行都不要
    trs = trs[1:41] + trs[42:50]

    dc_list = ["A668", "A672", "A673", "A680", "A681", "A715", "A716", "B194"]
    # 创建字典以存放库存数据
    result_dict = {}
    for d in dc_list:
        result_dict[d] = {}
        for c in code_list:
            result_dict[d][c] = {"dc": 0, "intransit": 0}

    for tr in trs:
        code = tr.find(".//td[3]/font/nobr").text.strip()                         # Material Code
        dc = tr.find(".//td[5]/font/nobr").text.strip()                           # DC
        unres_inv = int(float(tr.find(".//td[7]/font/nobr").text.strip()))        # Unrestrict Inventory
        qi_inv = int(float(tr.find(".//td[9]/font/nobr").text.strip()))           # Quality Inventory
        block_inv = int(float(tr.find(".//td[10]/font/nobr").text.strip()))       # Block Inventory
        intransit_inv = int(float(tr.find(".//td[13]/font/nobr").text.strip()))   # Intransit Inventory
        result_dict[dc][code]["dc"] = unres_inv + qi_inv + block_inv
        result_dict[dc][code]["intransit"] = intransit_inv

    for d in dc_list:
        result_string = str(result_dict[d][code_list[0]]["dc"]) + "," + str(result_dict[d][code_list[0]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[1]]["dc"]) + "," + str(result_dict[d][code_list[1]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[2]]["dc"]) + "," + str(result_dict[d][code_list[2]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[3]]["dc"]) + "," + str(result_dict[d][code_list[3]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[4]]["dc"]) + "," + str(result_dict[d][code_list[4]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[5]]["dc"]) + "," + str(result_dict[d][code_list[5]]["intransit"])
        print(result_string)

koala_code = []
get_data("E:/TDDOWNLOAD/ka.htm", koala_code)
