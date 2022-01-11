import time
import random
import requests
from openpyxl import load_workbook
import re
import json
import os
import csv






id = "11440600007039565U3440108001002"

# class Result_Page():
#     def __init__(self,id,item_name_array,example_GUID_array,form_GUID_array):
#         self.event_id = id
#         self.item_name_array = item_name_array
#         self.example_GUID_array = example_GUID_array
#         self.form_GUID_array = form_GUID_array
#
class One_log:
    def __init__(self,id,item_name,example_GUID,form_GUID,example_GUID_filename,form_GUID_filename,flage):
        self.event_id = id
        self.Item_name = item_name
        self.example_GUID = example_GUID
        self.form_GUID = form_GUID
        self.example_GUID_filename = example_GUID_filename
        self.form_GUID_filename = form_GUID_filename
        self.flage = flage
#
#



user_agent_list = ["Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36",
                    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36",
                    "Mozilla/5.0 (Windows NT 10.0; WOW64) Gecko/20100101 Firefox/61.0",
                    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36",
                    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.62 Safari/537.36",
                    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.101 Safari/537.36",
                    "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0)",
                    "Mozilla/5.0 (Macintosh; U; PPC Mac OS X 10.5; en-US; rv:1.9.2.15) Gecko/20110303 Firefox/3.6.15",
                    ]


def read_dict(dict):
    list = dict["AUDIT_MATERIAL"]
    log = []
    for item in list:

        example_len = len(item["EXAMPLE_GUID"])
        from_len = len(item["FORM_GUID"])

        if len(item["EXAMPLE_GUID"])>0 or len(item["FORM_GUID"])>0:
            material_nm = item["MATERIAL_NAME"].replace("\t","")
            download_path = id + "\\"+material_nm+"\\"

            check_path(download_path)


            try:
                sample = item["EXAMPLE_GUID"][0]
                sample_url = "https:"+sample["FILEPATH"]
                sample_path = download_path+"示例文件_"+sample["ATTACHNAME"]
                example_GUID_filename =  sample["ATTACHNAME"]
            except:
                sample = {}
                sample_url =""
                sample_path=""
                example_GUID_filename = ""

            try:
                blank_form = item["FORM_GUID"][0]
                blank_form_url = "https:" + blank_form["FILEPATH"]
                blank_form_path = download_path + "空白表格_" + blank_form["ATTACHNAME"]
                form_GUID_filename =  blank_form["ATTACHNAME"]

            except:
                blank_form = {}
                blank_form_url =""
                blank_form_path=""
                form_GUID_filename=""


            print(sample_url)
            # print(sample["ATTACHNAME"])
            # sample_type = get_file_type(sample)

            print(blank_form_url)
            # print(blank_form["ATTACHNAME"])
            # blank_form_type = get_file_type(blank_form)

            # one_log = One_log(id,material_nm,sample_url,blank_form_url,sample["ATTACHNAME"],blank_form["ATTACHNAME"],-1)

            One_log = {
                "event_id": id,
                "Item_name": material_nm,
                "example_GUID": sample_url,
                "form_GUID": blank_form_url,
                "example_GUID_filename": example_GUID_filename,
                "form_GUID_filename":form_GUID_filename,
                # "flage":-1
            }



            # 下载文件
            # download_one_file(sample_url,sample_path)
            # download_one_file(blank_form_url,blank_form_path)
            #
            # if sample_type==blank_form_type:
            #     sample_size = get_FileSize(sample_path)
            #     print(f"sample_size = {sample_size}")
            #     blank_form_size = get_FileSize(blank_form_path)
            #     print(f"blank_form_size = {blank_form_size}")
            #
            #     if sample_size==blank_form_size:
            #         One_log["flage"]=0
            #         print("error!")
            #
            #
            #
            #     else:
            #         One_log["flage"]=1
            #         print("success")
            # else:
            #     print("样本与空白表格类型不匹配，跳过")

            Save_log_json(One_log)

            log.append(One_log)

            # time.sleep(5)




            # print("sample = ",end="")
            # print(sample)

            # print("blank_form = ",end="")
            # print(blank_form)



            print("-----------")
        else:
            print("样本和表单为空")
            print(f"example = {example_len}")
            print(f"from = {from_len}")

    return log




def respose_by_id(id):
    url = "https://www.gdzwfw.gov.cn/portal/api/v2/item-event/getAuditItemDetailCur?TASK_CODE=" + id
    # 11440605007040232D
    headers = {'User-Agent': random.choice(user_agent_list)}
    response = requests.get(url, headers=headers).text
    # print(response)
    json1 = json.loads(response)
    # with open("json1.json",'w',encoding='utf-8')as f:
    #     r = json.dumps(json1, ensure_ascii=False,indent=4)
    #     f.write(r)
    #     f.close()
    return json1

def get_file_type(dict):
    name = dict["ATTACHNAME"]
    type = name.split(".")[1]
    print(type)
    return type

def download_one_file(url,path):
    headers = {'User-Agent': random.choice(user_agent_list)}
    req = requests.get(url,headers=headers)
    with open(path, "wb") as f:
            f.write(req.content)
            f.close()
    print("下载完成")

# def download_one_file_urllib2(url,path):
#     f = urllib2.urlopen(url)
#     data = f.read()
#     with open(path, "wb") as code:
#         code.write(data)



def get_FileSize(filePath):
    # filePath = unicode(filePath,'utf8')
    fsize = os.path.getsize(filePath)
    fsize = fsize/float(1024*1024)
    return fsize

def check_path(path):
    isExists = os.path.exists(path)
    if not isExists:
        os.makedirs(path)
        print(f"创建目录{path}")

def Save_log(Log):
    with open('Log.csv', 'a', encoding='utf-8', newline='')as f:
        writer = csv.DictWriter(f, fieldnames=['event_id', 'Item_name', 'example_GUID', 'form_GUID','example_GUID_filename','form_GUID_filename','flage'])
        writer.writeheader()
        for one_log in Log:
            writer.writerow(one_log)
    pass

# 先保存到json更方便
def Save_log_json(Log):
        with open("Log.txt", 'a', encoding='utf-8')as f:
            r = json.dumps(Log, ensure_ascii=False)
            f.write(r)
            f.write("\n")
            f.close()

def save_check_point(id):
    with open("check_point.txt","w")as f:
        f.write(id)
        f.close()

def load_check_point():
    with open("check_point.txt","r")as f:
        r = f.read()
        return r



xlsx_path = "test.xlsx"

if __name__ == '__main__':
    # download_by_url("https://static.gdzwfw.gov.cn/guide/files/c1a75a904cec495ca21446c5e970ace5","建筑物命名、更名审批表（样表）.doc")

    while 1:

        workbook = load_workbook(filename=xlsx_path)
        sheet = workbook["Sheet1"]
        ff = 0
        try:
            check_point = load_check_point()
        except:
            check_point = "11440600007039469J344014700200402"

        for cell in sheet['B']:
            id = cell.value

            if check_point!=id and ff==0:
                continue

            if check_point==id:
                ff=1


            print(id)
            if id=="实施编码":
                continue
            try:
                check_path(id)
                dict = respose_by_id(id)
                log = read_dict(dict)
                save_check_point(id)
            except:
                with open("error.txt", 'a')as f:
                    f.write(id)
                    f.write("\n")
                    f.close()
                continue
        break

    Save_log(log)

# with open("json1.json",'r',encoding='utf-8')as f:
#     read_str(f)
    # print(list)

    # print(str)
    # print(type(str))

