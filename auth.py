from flask import (
    Blueprint, flash, g, redirect, render_template, request, session, url_for,current_app
)
import random,time
import requests
import json
from urllib.parse import parse_qs
import uuid
def createPhoneCode(): 
    chars=['0','1','2','3','4','5','6','7','8','9'] 
    x = random.choice(chars),random.choice(chars),random.choice(chars),random.choice(chars) 
    verifyCode = "".join(x) 
    session["phoneVerifyCode"] = {"time":int(time.time()), "code":verifyCode} 
    return verifyCode
def sendPhoneMessage(phoneNumber = ''):
    url = 'https://smsapi.mitake.com.tw/api/mtk/SmSend'
    clientid = str(uuid.uuid4())  
    # phoneNumber = ''
    verifycode = createPhoneCode()
    smbody = f'FreeRcCheck的驗證碼為:{verifycode}'
    # url = 'https://smsapi.mitake.com.tw/api/mtk/SmQuery'
    params = {'CharsetURL': 'UTF-8',
                'username': '45008175SMS', 
                'password': "Elements25926882",
                'clinetid': clientid,
                'dstaddr':phoneNumber,
                'smbody':smbody,
            }

    response = requests.post(url, params=params)
    print(json.dumps(parse_qs(response.text)))

if __name__ == '__main__':
    print(str(uuid.uuid4()))