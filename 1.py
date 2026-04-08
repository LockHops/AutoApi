# -*- coding: UTF-8 -*-
import requests as req
import json, sys, time, os

path = sys.path[0] + r'/1.txt'
num1 = 0

CLIENT_ID = os.getenv("CONFIG_ID")
CLIENT_SECRET = os.getenv("CONFIG_KEY")


def gettoken(refresh_token):
    url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'

    data = {
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'redirect_uri': 'http://localhost:53682/'
    }

    res = req.post(url, data=data)
    jsontxt = res.json()

    print("TOKEN RESPONSE:", jsontxt)  # ⭐关键调试

    if 'refresh_token' not in jsontxt or 'access_token' not in jsontxt:
        raise Exception(f"Token refresh failed: {jsontxt}")

    new_refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']

    with open(path, 'w') as f:
        f.write(new_refresh_token)

    return access_token


def call_api(url, headers, name):
    global num1
    try:
        res = req.get(url, headers=headers)
        if res.status_code == 200:
            num1 += 1
            print(f"{name} 调用成功 {num1} 次")
        else:
            print(f"{name} 失败: {res.status_code} -> {res.text[:100]}")
    except Exception as e:
        print(f"{name} 异常: {e}")


def main():
    global num1

    with open(path, "r") as f:
        refresh_token = f.read().strip()

    localtime = time.asctime(time.localtime())

    access_token = gettoken(refresh_token)

    headers = {
        'Authorization': 'Bearer ' + access_token,  # ✅修复
        'Content-Type': 'application/json'
    }

    call_api('https://graph.microsoft.com/v1.0/me/drive/root', headers, "1")
    call_api('https://graph.microsoft.com/v1.0/me/drive', headers, "2")
    call_api('https://graph.microsoft.com/v1.0/drive/root', headers, "3")
    call_api('https://graph.microsoft.com/v1.0/users', headers, "4")  # ✅去掉空格
    call_api('https://graph.microsoft.com/v1.0/me/messages', headers, "5")
    call_api('https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules', headers, "6")
    call_api('https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta', headers, "7")
    call_api('https://graph.microsoft.com/v1.0/me/drive/root/children', headers, "8")
    call_api('https://api.powerbi.com/v1.0/myorg/apps', headers, "9")
    call_api('https://graph.microsoft.com/v1.0/me/mailFolders', headers, "10")

    print('此次运行结束时间为:', localtime)


for _ in range(7):
    main()
