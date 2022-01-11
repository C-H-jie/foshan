import time
import random
import requests
from openpyxl import load_workbook
import re
import json
import os
import csv




user_agent_list = ["Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36",
                    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36",
                    "Mozilla/5.0 (Windows NT 10.0; WOW64) Gecko/20100101 Firefox/61.0",
                    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36",
                    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.62 Safari/537.36",
                    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.101 Safari/537.36",
                    "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0)",
                    "Mozilla/5.0 (Macintosh; U; PPC Mac OS X 10.5; en-US; rv:1.9.2.15) Gecko/20110303 Firefox/3.6.15",
                    ]



def download_one_file(url,path):
    headers = {'User-Agent': random.choice(user_agent_list)}
    req = requests.get(url,headers=headers)
    with open(path, "wb") as f:
            f.write(req.content)
            f.close()
    print("下载完成")

def check_path(path):
    isExists = os.path.exists(path)
    if not isExists:
        os.makedirs(path)
        print(f"创建目录{path}")



if __name__ == '__main__':
    with open("Log.txt", "r", encoding="utf-8")as f:
        for line in f.readlines():
            one_log = json.loads(line)

            example_url = one_log["example_GUID"]
            form_url = one_log["form_GUID"]

            example_path= one_log["event_id"]+"\\"+one_log["Item_name"]+"\\"+"示例文件_"+one_log["example_GUID_filename"]
            form_path = one_log["event_id"] + "\\" + one_log["Item_name"] + "\\" + "空白表格_"+one_log["form_GUID_filename"]

            print(one_log)

            try:
                if len(example_url)>1:
                    download_one_file(example_url,example_path)
                if len(form_url)>1:
                    download_one_file(form_url,form_path)
            except:
                with open("download_error.txt","a",encoding="utf-8")as f:
                    f.write(line)
                    f.close()