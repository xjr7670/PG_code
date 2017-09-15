import os
import time
import json

import openpyxl
import requests
import xmltodict
from lxml import etree

class NYK_Track(object):
    """Track for NYK shipping container"""

    def __init__(self, session, *argv):

        self.timestamp = timestamp
        self.session = session
        self.url_list = argv

    
    def get_cntr_no(self):
        # Get the container number list

    
    def get_bkg_cop_no(self, cntr_no):
        # Get the bkg number and cop number

        self.data1 = {
            'cntr_no': cntr_no,
            'cust_cd': '',
            'f_cmd': 122
        }
        
        self.session.get(self.url_list[0])
        self.session.get(self.url_list[1])
        self.session.get(self.url_list[2])
        res = self.session.post(self.url_list[3], data=self.data1)
        res_json = res.json()
        bkg_no = res_json['list'][0]['bkgNo']
        cop_no = res_json['list'][0]['copNo']
        return (bkg_no, cop_no)

    def track(self, cntr_no):
        bkg_no, cop_no = self.get_bkg_cop_no()
        data2 = {
            'cntr_no': cntr_no,
            'bkg_no': bkg_no,
            'cop_no': cop_no,
            'f_cmd': 125
        }
        res = self.session.post(url3, data=self.data2)


if __name__ == "__main__":

    # General unix timestamp
    timestamp = int(round(time.time() * 1000))

    # General requests Session object with headers
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4,zh-TW;q=0.2',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Host': 'www.nykline.com',
        'X-Requested-With': 'XMLHttpRequest',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.101 Safari/537.36',
        'Referer': 'https://www.nykline.com/ecom/CUP_HOM_3301.do?sessLocale=en',
    }
    session = requests.Session(headers=headers)

    # Set url list
    url0 = "https://www.nykline.com/ecom/MenuGS.do?f_cmd=105&pagerows=&mnu_div_cd=E&hpg_lang_tp_cd=EN&user_id="
    url1 = "https://www.nykline.com/ecom/MenuGS.do?f_cmd=101&pagerows=&dsp_flg=&mnu_div_cd=F&hpg_lang_tp_cd=EN"
    url2 = "https://www.nykline.com/ecom/apps/gnoss/webservice/tracktrace/cargotracking/script/CUP_HOM_3301.js?version=1482308284000&_=%d" % timestamp
    url3 = "https://www.nykline.com/ecom/CUP_HOM_3301GS.do"
    url_list = [url0, url1, url2, url3]

    # Get container number list

    nyk_track = NYK_Track()