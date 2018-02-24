#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月10日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.GreenCaseJob.GreenCaseMonitor
@description: 绿网案例监控
'''
import sys
# import os
# sys.path.insert(0, os.path.abspath("../../"))
from collections import OrderedDict
from datetime import datetime
from email.mime.text import MIMEText
import re

from bs4 import BeautifulSoup
from jinja2 import Template
from openpyxl.reader.excel import load_workbook
import requests

from AutoReportUtil import sendEmail  # @UnresolvedImport
from jobs.GreenCaseJob.GreenCase import login  # @UnresolvedImport


# import AutoReportGlobals  # @UnresolvedImport
__version__ = "0.0.1"

# 普通工单
ReGeneralWorkOrder = re.compile("d\.add\((\d+),\d+,'([^\x00-\xff]+) \[([1-9]\d*)\]'")
# 危机工单
ReCrisisWorkOrder = re.compile("d_crisis\.add\((\d+),\d+,'([^\x00-\xff]+) \[([1-9]\d*)\]'")
# 升级工单
ReUpgradeWorkOrder = re.compile("d_upgrade\.add\((\d+),\d+,'([^\x00-\xff]+) \[([1-9]\d*)\]'")
# 已超时工单
ReOvertimeWorkOrder = re.compile("d_overtime\.add\((\d+),\d+,'([^\x00-\xff]+) \[([1-9]\d*)\]'")
# 预超时工单
RePretimeWorkOrder = re.compile("d_pretime\.add\((\d+),\d+,'([^\x00-\xff]+) \[([1-9]\d*)\]'")
# 工单量
ReCount = re.compile('<font color="red">(\d+)</font>')
# 简要模版
ReDetail = "\.add\(\d{11,20},%s,'\[(.*?)\](\d{11,20})','javascript"

# 工单数量地址
COUNT_URL = "http://10.109.214.54:8080/cscwf/wf/manage_support/ManageSupportInitAction.action?moduleID=M0004"
# 工单详情地址
DETAIL_URL = "http://10.109.214.54:8080/cscwf/wf/appealshowview/AppealShowAction.action?acceptId={0}&flag=1"
# 催单提醒
REMIND_URL = "http://10.109.214.54:8080/cscwf/wf/manage_support/ShowWfRemind.action?opFlag=manual"
HEADERS = {
    "Accept": "image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*",
    "Referer": "http://10.109.214.54:8080/cscwf/wf/manage_support/ManageSupportInitAction.action?moduleID=M0004",
    "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)",
    "Content-Type": "application/x-www-form-urlencoded",
    "Cache-Control": "no-cache"
}
BHEADERS = {
    "Accept": "image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*",
    "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)",
    "Content-Type": "application/x-www-form-urlencoded",
    "Cache-Control": "no-cache"
}
RHEADERS = {
    "Accept": "*/*",
    "Referer": "http://10.109.214.54:8080/cscwf/wf/manage_support/manage_support.jsp?module=M0004",
    "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)",
    "Content-Type": "application/x-www-form-urlencoded",
    "Cache-Control": "no-cache",
    "X-Requested-With": "XMLHttpRequest"
}
MailTemplate = Template("""<html>
    <body>
        <center><h1>{{ title }}</h1></center><br />
        <b>当前需处理的记录数为: <font color=red style="font-size:12px;font-weight:bold;">{{ count }}</font></b><br />
        {% for title, order in orders.items() %}
        <b>{{ title }}:</b><br />
            {% for item in order %}
            <b>　{{ item.get("name") }}: <font color=red style="font-size:12px;font-weight:bold;">{{ item.get("count") }}</font></b><br />
                {% for detail in item.get("details") %}
                |　　{{ loop.index }}. 工单(L{{ detail[0] }})：{{ detail[1] }}<br />
                |　　　{{ detail[2] }}
                {% endfor %}
            {% endfor %}
        {% endfor %}
        <b>工单催单:</b><br />
        |　{{ urges }}<br />
        <br />
        #说明 回复命令格式：<br />
        ##查看该工单的所有流转处理信息<br />
        L绿网工单号<br />
    <body>
</html>""")

        # # 绿网回复,工单号,0,回复内容　　　表示把工单转派到(0-眉山投诉组,1-省计费)<br />
        # # 绿网检索,50,检索内容　　　　　　　　表示从绿网案例库中搜索50条关键词案例(多个关键词用|分开)<br />

WxTemplate = Template("""{{ title }}
当前需处理的记录数为: {{ count }}
{% for title, order in orders.items() %}
  {{ title }}:
    {% for item in order %}
    {{ item.get("name") }}: {{ item.get("count") }}
        {% for detail in item.get("details") %}
        |　　{{ loop.index }}. 工单(L{{ detail[0] }})：{{ detail[1] }}
        {% endfor %}
    {% endfor %}
{% endfor %}
工单催单:
    |　{{ urges }}""")

class GreenCaseMonitor:
    
    def __init__(self, Id, tname, port):
        self.retrys = 0
        self.port = port
        self.logger = initLogger("datas/logs/" + Id, Id)
        wb = load_workbook(tname)
        # 预发送人
        preSheet = wb["预发送"]
        if not preSheet:
            self.logger.warn("无法找到预发送人表")
            self.preSend=[]
        else:
            self.preSend = [str(cell.value) for cell in preSheet["A"] if str(cell.value) != "None"]  # list(preSheet.columns)[0]]
        self.logger.debug("预发送人: {0}".format(self.preSend))
    
    def getLogin(self):
        while 1:
            self.retrys += 1
            self.cookies = login()  # 获取身份信息
            if not self.cookies:
                self.logger.error("获取身份信息失败,尝试次数: %s" % self.retrys)
            else:break
            if self.retrys == 5:
                raise Exception("获取身份信息失败")
            return sys.exit()
    
    def getItem(self, title, regx, texts):
        orders = re.findall(regx, texts)
        items = []
#         print(title,orders)
        self.logger.debug("%s工单:%s" % (title, orders))
        for code, name, count in orders:
            item = {"name":name, "count":count}
            details = []
#             print(ReDetail % code)
#             print(re.findall(ReDetail % code, texts))
            for brief, acceptId in re.findall(ReDetail % code, texts):
                req = requests.get(DETAIL_URL.format(acceptId), headers=BHEADERS, cookies=self.cookies, timeout=60)
                text = req.text.replace("\r\n", "#@#").replace("\n", "#@#")  # 替换换行
                bs = BeautifulSoup(text, "lxml")
                tds = bs.findAll("td", {"style":"word-break: break-all"})
                detail = str(tds[-1].getText(strip=True)) if tds else "\r\n"
#                 detail = detail.replace("<br>", "").replace("<br/ >", "").replace("\r\n", "<br/ >").replace("\n", "<br/ >")
                detail = detail.replace("#@#", "<br />").replace("\t", " ")
                if detail.startswith("<br />"):
                    detail = detail[6:]
                # print(detail)
                details.append((acceptId, brief, detail))
            item["details"] = details
            items.append(item)
        return items

    def start(self):
        self.getLogin()
        # 登录成功
        # 监控数据
        self.logger.debug("monitor start")
        req = requests.get(COUNT_URL, headers=HEADERS, cookies=self.cookies, timeout=60)
        if req.text.find("版权所有") > 0:  # 需要重新登录
            self.logger.debug("需要重新登录")
            self.getLogin()
        
        # 当前数量
        count = re.findall(ReCount, req.text)
        self.logger.debug("当前需处理的记录数为:%s" % count)
        
        # 有序字典
        orders = OrderedDict()
        
        # 普通工单
        gorders = self.getItem("普通工单", ReGeneralWorkOrder, req.text)
        # print("普通工单", gorders)
        self.logger.debug("普通工单:%s" % gorders)
        orders["普通工单"] = gorders
        
        # 危机工单
        corders = self.getItem("危机工单", ReCrisisWorkOrder, req.text)
        # print("危机工单", corders)
        self.logger.debug("危机工单:%s" % corders)
        orders["危机工单"] = corders
        
        # 升级工单
        uorders = self.getItem("升级工单", ReUpgradeWorkOrder, req.text)
        # print("升级工单", uorders)
        self.logger.debug("升级工单:%s" % uorders)
        orders["升级工单"] = uorders
        
        # 已超时工单
        oorders = self.getItem("已超时工单", ReOvertimeWorkOrder, req.text)
        # print("已超时工单", oorders)
        self.logger.debug("已超时工单:%s" % oorders)
        orders["已超时工单"] = oorders
        
        # 预超时工单
        porders = self.getItem("预超时工单", RePretimeWorkOrder, req.text)
        # print("预超时工单", porders)
        self.logger.debug("预超时工单:%s" % porders)
        orders["预超时工单"] = porders

        # 获取催单提醒
        urges = ""
        req = requests.get(REMIND_URL, headers=RHEADERS, cookies=self.cookies, timeout=60)
        if len(req.text) < 500:
            urges = req.text
        self.logger.debug("工单催单:%s" % urges)
        
        if all((not gorders, not corders, not uorders, not oorders, not porders)):
            # print("当前没有需要处理的工单")
            return self.logger.info("当前没有需要处理的工单")
        
        title = "{0} 绿网投诉统计".format(datetime.now().strftime("%Y/%m/%d %H:%M:%S"))
        text = MailTemplate.render(
            title=title,
            count=count[0] if count else "0",
            orders=orders,
            urges=urges.replace("<br/>", "").replace("<br>", "").replace("<br/ >", "").replace("\r\n", "<br/ >").replace("\n", "<br/ >")
        )
        # print(text)
        message = MIMEText(text, _subtype="html", _charset="gb2312")
        self.logger.debug(str(sendEmail(title, message, self.preSend)))
        self.logger.info("本时段绿网统计发送完毕")
#         if not AutoReportGlobals.WeixinBot:
#             return
        '''
        # 微信发送
        wxt = WxTemplate.render(
            title=title,
            count=count[0] if count else "0",
            orders=orders,
            urges=urges.replace("<br/>", "").replace("<br>", "").replace("<br/ >", "")
        )
        bs = BeautifulSoup(wxt, "html.parser")
        wxt = "".join(bs.stripped_strings)
#         print(wxt)
        _json = {
            "type":"text",
            "data":wxt
        }
        try:
            req = requests.post("http://127.0.0.1:%s/api/wechat/send" % self.port, json=_json)
            if req.json().get("code") != 1:
                return self.logger.error("微信群消息发送失败")
            self.logger.debug("微信群消息发送完毕")
        except Exception as e:
            self.logger.error(str(e))
        # AutoReportGlobals.WeixinBot.sendText(wxt)
        '''

if __name__ == "__main__":
    from AutoReportUtil import initLogger  # @UnresolvedImport
    g = GreenCaseMonitor("GreenCaseJob", "绿网监控配置文件.xlsx", 8989)
    g.start()
