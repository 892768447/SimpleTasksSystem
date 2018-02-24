#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月4日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.CommonLibs.EmailMonitor
@description: 
'''
# import sys, os
# sys.path.insert(0, os.path.abspath("../../"))
from concurrent.futures.thread import ThreadPoolExecutor
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import json
import re
import traceback

from bs4 import BeautifulSoup
from imapclient.imapclient import IMAPClient
import requests

import AutoReportGlobals  # @UnresolvedImport
from AutoReportUtil import sendMms, sendTextMms, sendEmail  # @UnresolvedImport


try: from jobs.GreenCaseJob.GreenCase import login  # @UnresolvedImport
except: pass


__version__ = "0.0.1"

PATTERN = re.compile(
"""(M[0-9a-zA-Z]{6}[,\d]+| #匹配彩信发送 M1000f9
    M[0-9a-zA-Z]{6}|       #匹配彩信发送 M1000f9,1,2,3,4,5,6,7,8,9,10,11
    T[0-9a-zA-Z]{6}[,\d]+| #匹配文本发送 T1000f9
    T[0-9a-zA-Z]{6}|       #匹配文本发送 T1000f9,1,2,3,4,5,6,7,8,9
    L\d{8,20})            #匹配绿网工单 L2017070525624
""", re.X)
EMAILPATTERN = re.compile("\<(\d+)@139\.com\>")  # re.compile("\<(\d+@139\.com)\>")
# 工单详情地址
DETAIL_URL = "http://10.109.214.54:8080/cscwf/wf/appealshowview/AppealShowAction.action?acceptId={0}&flag=1"
HOST = "imap.139.com"
PORT = 143
# USERNAME = "135080*****@139.com"
# PASSWORD = "1qaz2wsx"
BHEADERS = {
    "Accept": "image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*",
    "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)",
    "Content-Type": "application/x-www-form-urlencoded",
    "Cache-Control": "no-cache"
}

class EmailMonitor:

    def __init__(self, logger):
        config = json.load(open("jobs/EmailMonitorJob/config.json", "r"))
#         print(config)
        self.USERNAME = config.get("USERNAME", None)
        self.PASSWORD = config.get("PASSWORD", None)
#         print(self.USERNAME)
        if not self.USERNAME or not self.PASSWORD:
            logger.error("can not find username and password")
            raise Exception("can not find username and password")
        self.isLogin = False
        self.logger = logger
        self.threadPool = ThreadPoolExecutor(max_workers=5)  # 线程池
    
    def _login(self):
        try:
            self.logger.debug("尝试登录邮箱")
            self.client = IMAPClient(HOST, PORT)
            loginInfo = self.client.login(self.USERNAME, self.PASSWORD)
            self.logger.debug("loginInfo: {0}".format(loginInfo))
            self.logger.debug("邮箱登录成功")
            self.isLogin = True
        except Exception as e:
            self.isLogin = False
            self.logger.error("邮箱登录失败: " + str(e))

    def start(self):
        # 先登录
        if not self.isLogin:
            self._login()
        if not self.isLogin:
            return self.logger.warn("邮箱未登录")
        try:
            noopInfo = self.client.noop()  # 发送心跳
            self.logger.debug("noopInfo: {0}".format(noopInfo))
            self.getEmails()
        except Exception as e:
            self.isLogin = False
            self.logger.error("".join(("获取邮件列表失败: ",
                                      str(e), "\n", traceback.format_exc())))

    def getEmails(self):
        # 进入收件箱
        self.client.select_folder("INBOX", False)
        # 获取未读邮件
        unseens = self.client.search("UNSEEN", "gb2312")
        msgDict = self.client.fetch(unseens, ["BODY.PEEK[]"])
        for msgid, message in msgDict.items():
            # 邮件内容
            msg = email.message_from_bytes(message[b"BODY[]"])  # 发送人
            From = email.header.make_header(email.header.decode_header(msg["From"]))
            From = EMAILPATTERN.findall(str(From))
#             print("From: ", From)
            # 邮件编码
            encoding = msg.get_content_charset()
            if not encoding:
                encoding = [en for en in msg.get_charsets() if en]
                encoding = "gb2312" if not encoding else encoding[0]
            self.logger.debug("email encoding: {0}".format(encoding))
            mainType = msg.get_content_maintype()
            mail_content = ""
            if mainType == "multipart":
                for part in msg.get_payload():
                    if part.get_content_maintype() == "text":
                        mail_content += part.get_payload(
                            decode=True).decode(encoding, errors="ignore")
            elif mainType == "text":
                mail_content = msg.get_payload(
                    decode=True).decode(encoding, errors="ignore")
            bs = BeautifulSoup(mail_content, "lxml")
            mail_content = ""
            for _content in bs.strings:  # 提炼所有文字（包含空格）
                mail_content += _content
            # 切割
            mail_content = mail_content[0:mail_content.find("#说明")]
            # 标记邮件已读
            self.client.set_flags(msgid, "SEEN")
            try:
                self.parseEmail(From, mail_content)
            except Exception as e:
                traceback.print_exc()
                self.logger.warn("parseEmail error: {0}".format(str(e)))

    def parseEmail(self, From, mail_content):
#             self.logger.debug("mail_content: {0}".format(mail_content))
        results = PATTERN.findall(mail_content)
        if not results:
            return self.logger.warn("邮件内容没有匹配到结果")
        results = results[0]
        ver, other = results[0], results[1:]
        other = re.split(",|，", other)  # 正则分割
        eid, whichs = other[0], other[1:]
        if ver == "M":
            self.doMmsMail(From, eid, whichs)  # 处理图片彩信
        elif ver == "T":
            self.doTextEmail(From, eid)  # 处理文字彩信
        elif ver == "L":
            self.doGreenCase(From, eid)  # 绿网查看
    
    def doTextEmail(self, From, eid):
        self.logger.debug("doTextEmail From:%s eid: %s" % (From, eid))
        if not From or not eid:
            return self.logger.warn("发送人或者ID未找到")
        self.logger.debug("email monitor :%s %s" % (eid, AutoReportGlobals.EmailIds.get(eid)))
        emailInfo = AutoReportGlobals.EmailIds.get(eid)
        if not emailInfo:
            def errorReply(eid, From):
                # 发送邮件
                try:
                    message = MIMEMultipart()
                    subject = "控制-错误"
                    message.attach(MIMEText("未找到该email id: {0}".format(eid)))
                    sendEmail(subject, message, From)
                    del AutoReportGlobals.EmailIds[eid]  # 删除
                except Exception as e:
                    self.logger.error(str(e))
            self.threadPool.submit(errorReply, eid, From)
            return self.logger.error("未找到该email id: {0}".format(eid))
        subject = emailInfo.get("subject")  # 获取主题
        content = emailInfo.get("content")  # 
        peoples = emailInfo.get("formalSend")  # 获取正式发送人

        def controlReply(subject, content, peoples, From):
            # 发送文本彩信
            try:
                result = sendTextMms(subject, content, peoples)
                message = MIMEMultipart()
                subject = "控制-" + ("成功" if result == 1 else "失败")
                message.attach(MIMEText(str(result)))
                sendEmail(subject, message, From)
                del AutoReportGlobals.EmailIds[eid]  # 删除
            except Exception as e:
                self.logger.error(str(e))
            del AutoReportGlobals.EmailIds[eid]  # 删除
        return self.threadPool.submit(controlReply, subject, content, peoples, From)

    def doMmsMail(self, From, eid, whichs):
        self.logger.debug("doMmsMail From:%s, eid: %s, whichs: %s" % (From, eid, whichs))
        if not From or not eid:
            return self.logger.warn("发送人或者ID未找到")

        whichs = [w for w in whichs if w != ""]  # 去掉空
#             import os
#             print(os.getpid())
#             print("email monitor :", eid, AutoReportGlobals.EmailIds.get(eid), AutoReportGlobals)
        self.logger.debug("email monitor :%s %s" % (eid, AutoReportGlobals.EmailIds.get(eid)))
#             print(id(AutoReportGlobals.EmailIds))
        emailInfo = AutoReportGlobals.EmailIds.get(eid)
#                 print("ejob: ", AutoReportGlobals.EmailIds)
        if not emailInfo:
            def errorReply(eid, From):
                # 发送邮件
                try:
                    message = MIMEMultipart()
                    subject = "日报控制-错误"
                    message.attach(MIMEText("未找到该email id: {0}".format(eid)))
                    sendEmail(subject, message, From)
                    del AutoReportGlobals.EmailIds[eid]  # 删除
                except Exception as e:
                    self.logger.error(str(e))
            self.threadPool.submit(errorReply, eid, From)
            return self.logger.error("未找到该email id: {0}".format(eid))
        subject = emailInfo.get("subject")  # 获取主题
        dataDir = emailInfo.get("dataDir")  # 获取数据目录
        peoples = emailInfo.get("formalSend")  # 获取正式发送人

        self.logger.debug("eid:{0}, subject:{1}, dataDir:{2}, peoples:{3}, whichs:{4}".format(eid, subject, dataDir, peoples, whichs))
        def controlReply(subject, dataDir, peoples, whichs, From):
            # 发送彩信和邮件
            try:
                result = sendMms(subject, dataDir, peoples, whichs, "")
                message = MIMEMultipart()
                subject = "日报控制-" + ("成功" if result == 1 else "失败")
                message.attach(MIMEText(str(result)))
                sendEmail(subject, message, From)
                del AutoReportGlobals.EmailIds[eid]  # 删除
            except Exception as e:
                self.logger.error(str(e))
            del AutoReportGlobals.EmailIds[eid]  # 删除
        return self.threadPool.submit(controlReply, subject, dataDir, peoples, whichs, From)
    
    def doGreenCase(self, From, eid):
        self.logger.debug("doGreenCase From:%s eid: %s" % (From, eid))
        if not From or not eid:
            return self.logger.warn("发送人或者ID未找到")
        def caseView(From, eid):
            try:
                req = requests.get(DETAIL_URL.format(eid), headers=BHEADERS,
                    cookies=AutoReportGlobals.GreenCookie, timeout=60)
                if req.text.find("版权所有") > 0:  # 需要重新登录
                    self.logger.warn("需要重新登录")
                    login()  # 重新登录
                    self.logger.debug("登录成功")
                    req = requests.get(DETAIL_URL.format(eid), headers=BHEADERS,
                        cookies=AutoReportGlobals.GreenCookie, timeout=60)
#                             print(req, req.text.find("版权所有"))
                    if req.text.find("版权所有") > 0:  # 需要重新登录
                        return self.logger.warn("又需要重新登录,放弃此操作")
                if not hasattr(self, "CaseStyle"):  # ../../datas/jobs/EmailMonitorJob/
                    self.CaseStyle = open("jobs/EmailMonitorJob/CaseStyle.css", "rb").read().decode() + "</style>"
                html = req.text.replace('class="menu_title"', 'class="menu_title" style="display: none;"').replace("</style>", self.CaseStyle)  # 替换样式表
                self.logger.debug("email len: %s" % len(html))
                # 发送邮件
                message = MIMEMultipart()
                subject = "工单查看-%s" % eid
                message.attach(MIMEText(html, _subtype="html", _charset="gb2312"))
                sendEmail(subject, message, From)
                self.logger.debug("工单查看邮件已发送")
            except Exception as e:
                self.logger.error(str(e))
#                 caseView(code, From)  # 调试
        self.threadPool.submit(caseView, From, eid)
        self.logger.debug("绿网查看邮件后台发送中")

def test():
    COMMANDS = """
    # M1000f9
    # M1000f9,1
    # M1000f9,1,2，3,4,5,6,7,8,9,10,11
    # T1000f9，
    # T1000f9
    # T1000f9,,,，，，,
    # T1000f9,1,2,3,4，5,6,7,8,9
    # L2017070525624
    """
    PATTERN = re.compile(
        """(M[0-9a-zA-Z]{6}(?:[,|，]\d)+|  #匹配彩信发送 M1000f9
            M[0-9a-zA-Z]{6}|          #匹配彩信发送 M1000f9,1,2,3,4,5,6,7,8,9,10,11
            T[0-9a-zA-Z]{6}(?:[,|，]\d)+|   #匹配文本发送 T1000f9
            T[0-9a-zA-Z]{6}|          #匹配文本发送 T1000f9,1,2,3,4,5,6,7,8,9
            L\d{8,20})                #匹配绿网工单 L2017070525624
        """, re.X)
    for item in PATTERN.findall(COMMANDS):
        print("item: ", item)

if __name__ == "__main__":
    test()
#     from AutoReportUtil import initLogger  # @UnresolvedImport
#     logger = initLogger("../../datas/logs/EmailMonitorJob/", "EmailMonitorJob")
#     g = EmailMonitor(logger)
#     g.start()
#     print("end")
