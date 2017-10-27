import os
from lxml import etree

os.chdir(r"C:/users/xian.jr/Documents/SAP/SAP GUI/")

def get_data(fname, code_list):
    """接受文件名和Code列表来处理数据"""
    
    src = open(fname).read()
    htm = etree.HTML(src)
    p = htm.xpath("//blockquote/p[2]")[0]
    trs = p.findall(".//tr")
    # 开头行是表头，结尾行是表尾，中间还有位于第2页的表头，是第41个tr
    # 这3行都不要
    trs = trs[1:41] + trs[42:82]

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
        unres_inv = int(float(tr.find(".//td[7]/font/nobr").text.strip().replace(",", "")))        # Unrestrict Inventory
        qi_inv = int(float(tr.find(".//td[9]/font/nobr").text.strip().replace(",", "")))           # Quality Inventory
        block_inv = int(float(tr.find(".//td[10]/font/nobr").text.strip().replace(",", "")))       # Block Inventory
        intransit_inv = int(float(tr.find(".//td[13]/font/nobr").text.strip().replace(",", "")))   # Intransit Inventory
        result_dict[dc][code]["dc"] = unres_inv + qi_inv + block_inv
        result_dict[dc][code]["intransit"] = intransit_inv

    for d in dc_list:
        result_string = str(result_dict[d][code_list[0]]["dc"]) + "," + str(result_dict[d][code_list[0]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[1]]["dc"]) + "," + str(result_dict[d][code_list[1]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[2]]["dc"]) + "," + str(result_dict[d][code_list[2]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[3]]["dc"]) + "," + str(result_dict[d][code_list[3]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[4]]["dc"]) + "," + str(result_dict[d][code_list[4]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[5]]["dc"]) + "," + str(result_dict[d][code_list[5]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[6]]["dc"]) + "," + str(result_dict[d][code_list[6]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[7]]["dc"]) + "," + str(result_dict[d][code_list[7]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[8]]["dc"]) + "," + str(result_dict[d][code_list[8]]["intransit"]) + ","
        result_string += str(result_dict[d][code_list[9]]["dc"]) + "," + str(result_dict[d][code_list[9]]["intransit"]) + ","
        print(result_string)

koala_code = ["82261753", "82261756", "82261762", "82261951", "82263780", "82264483"]
tampax_code = ["82255634", "82258740", "82259122", "82258746", "82258747", "82261076", "82273315", "82273316", "82277425", "82277427"]
# get_data("ka.htm", koala_code)
get_data("ta.htm", tampax_code)