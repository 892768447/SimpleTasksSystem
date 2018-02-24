#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月10日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.tasks.GreenCase
@description: 绿网案例监控和入库
'''
import hmac
import re

import requests

import AutoReportGlobals  # @UnresolvedImport


__version__ = "0.0.1"

key_pattern = re.compile("hex_hmac_md5\('(.*?)'")
token_pattern = re.compile('innerToken.value = "(.*?)";')
jsessionid_pattern = re.compile('JSESSIONID.value = "(.*?)";')

headers = {
    "Accept":"image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*",
    "Referer":"http://10.109.214.54:8080/csc/",
    "User-Agent":"Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)",
    "Content-Type": "application/x-www-form-urlencoded",
    "Cache-Control": "no-cache"
}

def getKey(html):
    try:
        return re.findall(key_pattern, html)[0].encode()
    except:
        return "3W-8kNVBTUt43ZoeXo-9hgF".encode()

def enPwd(key, pwd):
    myhmac = hmac.new(key)
    myhmac.update(pwd.encode())
    pwd = myhmac.hexdigest()
    # print("pwd", pwd)
    return pwd

def getInnerToken(html):
    try:
        return re.findall(token_pattern, html)[0]
    except:
        return ""

def getJsessionid(html):
    try:
        return re.findall(jsessionid_pattern, html)[0]
    except:
        return ""

def login():
    #*******第1步,获取首页,获取cookie和hex_hmac_md5_key
    req = requests.get("http://10.109.214.54:8080/csc/")
    JSESSIONID = req.cookies.get("JSESSIONID")
    # print("Cookie: JSESSIONID", JSESSIONID)  # 0000SSqozV07Cy3oOkKL6iyX42d:14umcpq5k
    # 获取hex_hmac_md5 key#SSqozV07Cy3oOkKL6iyX42d
    hex_hmac_md5_key = getKey(req.text)
    # print("hex_hmac_md5_key", hex_hmac_md5_key)
    
    #*******第2步,提交登录数据,获取cookie
    cookies = {
        "JSESSIONID":JSESSIONID
    }
    
    # 第2步登录,获取token
    data = {
        "j_username":"MSJF001",
        "j_password":enPwd(hex_hmac_md5_key, "ZZ@@2wsx3edc"),
        "j_recivePassword":"",
        "b_recivePassword":""
    }
    req = requests.post(
        "http://10.109.214.54:8080/csc/login.action",
        data=data,
        cookies=cookies,
        headers=headers
    )
    if req.text.find("goOn()") > -1:  # 登录成功
        innerToken = getInnerToken(req.text)
        JSESSIONID = getJsessionid(req.text)
        # print("innerToken", innerToken)  # C33C4157A804D17A71C7B1D3C695471A
        # print("JSESSIONID", JSESSIONID)  # SSqozV07Cy3oOkKL6iyX42d
        if not innerToken or not JSESSIONID:
            return  # print("登录失败：无法找到token")
    
    # 3.登录LoginNewServlet
    data = {
        "innerToken":innerToken,
        "JSESSIONID":JSESSIONID
    }
    headers["Referer"] = "http://10.109.214.54:8080/csc/login.action"
    req = requests.post(
        "http://10.109.214.54:8080/cscwf/LoginNewServlet",
        data=data,
        cookies=cookies,
        headers=headers,
        allow_redirects=False
    )
    # print(req.headers)
    # Location: http://10.109.214.54:8080/cscwf/j_security_check?j_username=MSJF001&j_password=ZZ@@2wsx3edc
    Location = req.headers.get("Location")
    # print("Location: ", Location)
    if Location.find("j_security_check") == -1:
        return None
    
    # 4.获取cookie
    req = requests.get(Location, cookies=cookies, headers=headers, allow_redirects=False)
    LtpaToken = req.cookies.get("LtpaToken")
    LtpaToken2 = req.cookies.get("LtpaToken2")
    # print("LtpaToken:", LtpaToken)
    # print("LtpaToken2:", LtpaToken2)
    cookies["LtpaToken"] = LtpaToken
    cookies["LtpaToken2"] = LtpaToken2
    # 登录成功
    AutoReportGlobals.GreenCookie = cookies
    return cookies
