from os import error, path
import sys
from typing import Text
import uuid
import requests
import argparse
import hashlib
import time
import pandas as pd
import json
import numpy as np
import os
import random

# from dotenv import load_dotenv
# dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
# if os.path.exists(dotenv_path):
#     load_dotenv(dotenv_path)

YOUDAO_URL=os.environ.get("YOUDAO_URL") or None
YOUDAO_APP_KEY=os.environ.get("YOUDAO_APP_KEY") or None
YOUDAO_APP_SECRET=os.environ.get("YOUDAO_APP_SECRET") or None
# from imp import reload

EXCEL_IN_PATH = './sys_reiview_1-2200_Jiaxin_1.xlsx'

EXCEL_OUT_PATH = './sys_reiview_1-2200_Jiaxin_1_out.xlsx'



class YOUDAO_Translator(object):
    def __init__(self, url, app_key, app_secret):
        self.YOUDAO_URL = url
        self.APP_KEY = app_key
        self.APP_SECRET = app_secret

    def encrypt(self, signStr):
        hash_algorithm = hashlib.sha256()
        hash_algorithm.update(signStr.encode('utf-8'))
        return hash_algorithm.hexdigest()

    def truncate(self, q):
        if q is None:
            return None
        size = len(q)
        return q if size <= 20 else q[0:10] + str(size) + q[size - 10:size]


    def do_request(self, url, data, try_num=5):
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        user_agent_list = ["Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36",
                    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36",
                    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36",
                    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.62 Safari/537.36",
                    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.101 Safari/537.36",
                    "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0)",
                    "Mozilla/5.0 (Macintosh; U; PPC Mac OS X 10.5; en-US; rv:1.9.2.15) Gecko/20110303 Firefox/3.6.15",]
        # print(random.choice(user_agent_list))
        user_agent = random.choice(user_agent_list)
        # print(user_agent)
        headers['User-Agent'] = user_agent
        while try_num > 0:
            try:
                result = requests.post(url, data=data, headers=headers)
                try_num = 0
            except requests.exceptions.ConnectionError as ex:
                if try_num <= 0:
                    print("Failed to retrieve: " + url + "\n" + str(ex))
                else:
                    print("{}:连接失败，剩余尝试次数:{}".format(url, try_num))
                    try_num -= 1
                    time.sleep(0.5)

        return result

    def connect(self, q, from_lan='en', to_lan='zh-CHS'):
        data = {}
        data['from'] = from_lan
        data['to'] = to_lan
        data['signType'] = 'v3'
        curtime = str(int(time.time()))
        data['curtime'] = curtime
        salt = str(uuid.uuid1())
        signStr = self.APP_KEY + self.truncate(q) + salt + curtime + self.APP_SECRET
        sign = self.encrypt(signStr)
        data['appKey'] = self.APP_KEY
        data['q'] = q
        data['salt'] = salt
        data['sign'] = sign
        # data['vocabId'] = "您的用户词表ID"

        response = self.do_request(self.YOUDAO_URL, data)
        contentType = response.headers['Content-Type']
        if contentType == "audio/mp3":
            millis = int(round(time.time() * 1000))
            filePath = "合成的音频存储路径" + str(millis) + ".mp3"
            fo = open(filePath, 'wb')
            fo.write(response.content)
            fo.close()
        else:
            # print(response.content)
            return response.content
    
    def tanslator(self, text, from_lan='en', to_lan='zh-CHS'):
        text_translated_bytes = self.connect(text, from_lan=from_lan, to_lan=to_lan)
        text_translated_str = text_translated_bytes.decode("utf-8")
        text_translated_dict = json.loads(text_translated_str)
        # print(text_translated_dict)
        if text_translated_dict['errorCode'] == "0":
            return text_translated_dict['translation']
        else:
            return None



def load_excel(file_path=EXCEL_IN_PATH, start_index=0):
    data = pd.read_excel(file_path, sheet_name=0)
    data_selected = data[data['#'] >= start_index]
    return data_selected




if __name__ == "__main__":
    START_INDEX = 0
    
    if os.path.exists(EXCEL_OUT_PATH):
        data_out = load_excel(file_path=EXCEL_OUT_PATH)
    else:
        data_selected = load_excel(start_index=0)
        data_out = data_selected.copy()
        data_out['']=np.nan
        data_out['Article Title CN']=' '
        data_out.to_excel(EXCEL_OUT_PATH, index=False)

    # print(data_out.head())
    
    # sys.exit(0)

    YOUDAO  = YOUDAO_Translator(YOUDAO_URL, YOUDAO_APP_KEY, YOUDAO_APP_SECRET)
    # data_selected.index()÷
    # data_out = pd.DataFrame()
    count = 0
    for idx, item in data_out.iterrows():
        if item['#'] >=START_INDEX:
            # if not np.isnan(item['Article Title CN']):
            if item['Article Title CN'] != ' ':
                continue
            text_en = item['Article Title']

            text_trans_array = YOUDAO.tanslator(text_en)

            text_translated = ''
            if text_trans_array is None:
                pass
            else:
                text_translated=''
                for text_idx, txt in enumerate(text_trans_array):
                    text_translated += txt
                    if text_idx < len(text_trans_array)-1:
                        text_trans_array += ";"
            print(idx, text_translated)
            data_out.loc[idx, 'Article Title CN'] = text_translated
            data_out.to_excel(EXCEL_OUT_PATH, index=False)
            time.sleep(0.5) #sleep一下，以防服务器将请求认为是攻击

    print("转译完成")
    data_out.to_excel(EXCEL_OUT_PATH, index=False)
