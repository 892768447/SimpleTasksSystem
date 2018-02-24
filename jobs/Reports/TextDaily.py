#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年8月25日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.Reports.TextDaily
@description: 文本彩信公共模块
'''

__version__ = "0.0.1"

from datetime import datetime
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from time import localtime, strftime

from jinja2 import Template
from openpyxl.reader.excel import load_workbook
import requests
from sqlalchemy.engine import create_engine

from AutoReportUtil import sendEmail, makeId  # @UnresolvedImport
# @UnresolvedImport
from AutoReportUtil import initLogger, DATE_FORMAT, CURRENT_TIME, CURRENT_MONTH


# import AutoReportGlobals  # @UnresolvedImport
__version__ = "0.0.1"

URL = "http://10.105.4.50:58080/BingoInsight/dataservice.dsr"

HEADERS = {
    "Accept": "*/*",
    "Referer": "http://10.105.4.50:58080/BingoInsight/portal/ClientBin/Bingosoft.DataOne.Portals.Boot.xap",
    "Content-Type": "text/json;chartset=utf-8",
    "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; InfoPath.2)",
    "Connection": "Keep-Alive",
    "Cache-Control": "no-cache",
    "Cookie": "JSESSIONID=D639101A895F1F7DC61011877347D8AC",
}

MailTemplate = Template("""<html>
    <body>
        <center><h1>{{ subject }}</h1></center><br />
        {{ content }}
        <br />
        #说明 任务ID：{{ eid }}<br />
        #回复命令格式：(注意回复不要附带原始邮件内容)<br /><br />
        ##正式发送：<br />
        T{{ eid }}<br />
    <body>
</html>""")


class TextDaily:

    def __init__(self, Id, config, path, subject, port=8989):
        os.makedirs("datas/logs/", exist_ok=True)
        os.makedirs("datas/tmps/", exist_ok=True)
        self.logger = initLogger("datas/logs/" + Id, Id)
        self.config = config
        self.subject = subject
        self.port = port
        filename = os.path.basename(path)  # 文件名
        self.filein = path
        self.fileout = os.path.join("datas/tmps/" + filename)
        self.dataDir = os.path.join(
            "datas/tmps/", os.path.basename(os.path.splitext(self.fileout)[0]))
        os.makedirs(self.dataDir, exist_ok=True)

    def start(self):
        self.logger.debug("pid: %s" % os.getpid())
        # 加载模版文件开始任务
        wb = load_workbook(self.filein)
        # 解析配置
        self.parseConfig(wb)
        # 预发送邮件
        if not self.preSend:
            return self.logger.warn("预发送人员列表为空")
        # 入库数据
        self.fillData(wb)
        # 查询结果
        content = self.getQueryMsg()
        eid = makeId()  # 获取唯一编码
        einfo = {
            "typ": 0,  # 储存方式,0=预发送,1=正式发送(邮件方式)
            "subject": self.subject,  # 储存主题信息
            "content": content,  #
            "preSend": self.preSend,  # 储存预发送人
            "formalSend": self.formalSend  # 储存正式发送人
        }
        # 此处改为post本地提交
#         AutoReportGlobals.EmailIds[eid] = einfo
        # 此处改为post本地提交
        _json = {
            "type": 0,  # 储存id
            "eid": eid,
            "einfo": einfo
        }
        req = requests.post("http://127.0.0.1:%s/api/email" %
                            self.port, json=_json)
        if req.json().get("code") != 1:
            return self.logger.error("储存生成的email id 错误")
        self.logger.debug("储存生成的email id 成功")
#         print("sqdjob: ", AutoReportGlobals.EmailIds)
        self._sendEmail(eid, self.subject, content,
                        self.preSend, self.formalSend, 0)
        self.logger.info("send pre email success")

    def _sendEmail(self, eid, subject, content, preSend, formalSend, typ=0):
        peoples = formalSend if typ else preSend  # 如果采用邮箱正式发送
        subject = subject + datetime.now().strftime(" - %Y%m%d %H:%M:%S")
        subject = subject if typ else subject + " - 预览"
        message = MIMEMultipart()
        message["Subject"] = Header(subject, "gb2312")  # 邮件主题
        # 邮件正文
        msg = MIMEText(MailTemplate.render(
            subject=subject, content=content, eid=eid
        ), _subtype="html", _charset="gb2312")
        message.attach(msg)
#         content = MIMEText(content + "\n\n控制命令：   T%s" % eid, _charset="gb2312")
#         message.attach(content)
        sendEmail(subject, message, peoples)

    def parseConfig(self, wb):
        # 发送人配置
        if not self.config.get("文本", None):
            raise Exception("无法找到文本发送配置")
        # 预发送人
        preSheet = wb["预发送"]
        if not preSheet:
            raise Exception("无法找到预发送人表")
        self.preSend = [str(cell.value) for cell in preSheet["A"] if str(
            cell.value) != "None"]  # list(preSheet.columns)[0]]
        self.logger.debug("预发送人: {0}".format(self.preSend))
        # 获取正式发送人
        formalSheet = wb["联系人"]
        if not formalSheet:
            raise Exception("无法找到联系人表")
        self.formalSend = [str(cell.value) for cell in formalSheet["A"] if str(
            cell.value) != "None"]  # list(formalSheet.columns)[0]]
        self.logger.debug("正式发送人: {0}".format(self.formalSend))
        # 填充区域配置
        self.fillArea = self.config.get("入库", [])
#         if not self.fillArea:
#             raise Exception("入库配置错误")
        self.logger.debug("end parse config")

    def getQueryMsg(self):
        text = self.config.get("文本", {})
        dburl = text.get("类型", "")
        query = text.get("语句", "")
        if not dburl or not query:
            raise Exception("文本配置有误")
        return self.queryDb(dburl, query)[0][0]

    def fillData(self, wb):
        self.logger.debug("start fill data")
#         for sheetName, items in self.fillArea.items():
        for item in self.fillArea:
            self.fillItem(wb, item)
        self.logger.debug("end fill data")

    def fillItem(self, wb, item):
        isClear = item.get("清空", "否") == "是"
        typ = item.get("类型", "")
        dbcon = item.get("数据库", "")
        dbsrc = item.get("数据源", "")
        tname = item.get("表名", "")
        query = item.get("语句", "")
        self.logger.debug("isClear:{0}\n"
                          "dbcon:{1}\n"
                          "dbsrc: {2}\n"
                          "tname:{3}\n"
                          "type:{4}\n".format(isClear,
                                              dbcon, dbsrc, tname, typ
                                              )
                          )
        if not dbcon:
            return self.logger.warn("入库数据库未配置")
        if not tname:
            return self.logger.warn("未配置表名")
        if not query:
            return self.logger.warn("查询语句不能为空")
        self.logger.debug("start get data")
        if typ == "selfsql":
            # 自助报表
            times, datas = self.queryWeb(dbsrc, query)
        else:
            # 连接数据查询
            times, datas = set(), self.queryDb(typ, query)
        self.logger.debug("end get data")

        self.logger.debug("start fill data")
        if not datas:
            return self.logger.debug("end fill data: data is empty")
        _cols = []
        for r, item in enumerate(datas):  # @UnusedVariable
            _col = []
            for c, value in enumerate(item):
                if isinstance(value, str) and 1 <= len(value) <= 8:
                    try:
                        value = float(value)  # 尝试转成float类型
                    except:
                        pass
                try:  # 时间戳转换
                    if c in times:
                        value = localtime(int(str(value)[:10]))
                        value = strftime("%Y-%m-%d %H:%M:%S", value)
                except Exception as e:
                    self.logger.warn(str(e))
                _col.append(value)
            _cols.append(_col)
        engine = create_engine(dbcon)
        if isClear:
            #             print("truncate table %s" % tname)
            engine.execute("truncate table %s" % tname)  # 清空表数据
        with engine.begin() as connection:
            keys = ",".join(["%s" for _ in datas[0]])
            for data in _cols:
                #                 rcount = engine.execute("insert into {tname} values ({values})".format(
                # tname=tname, values=",".join(["'%s'"%d for d in data])))
                connection.execute("insert into {tname} values ({keys})".format(
                    tname=tname, keys=keys), data)
#             connection.commit()
        self.logger.debug("end fill %s row data" % len(datas))

    def queryWeb(self, dbsrc, query):
        # 时间戳索引记录
        times = set()
        if not dbsrc:
            self.logger.error("未选择数据源")
            return times, []
        query = Template(query).render(
            DATE_FORMAT=DATE_FORMAT, CURRENT_MONTH=CURRENT_MONTH, CURRENT_TIME=CURRENT_TIME)
        self.logger.debug("queryWeb: query: %s" % query)
        data = {
            "CommandName": "metadata.queryWithDS",
            "Params": {
                "querySQL": query + " FETCH FIRST 1000 ROWS ONLY",
                "dataSourceName": dbsrc
            }
        }
        req = requests.post(URL, json=data, headers=HEADERS)
        result = req.json()
        resultCode = result.get("resultCode", 0)
        resultDesc = result.get("resultDesc", "未知错误")
        self.logger.debug("resultCode: {0}".format(resultCode))
        self.logger.debug("resultDesc: {0}".format(resultDesc))
        if resultCode != 200:
            return times, []

        resultValue = result.get("resultValue", {})
        # 表头
        headers = resultValue.get("columnNames", [])
        for col, cvalue in enumerate(headers):  # 写入表头
            #             print(col,cvalue,cvalue.lower().find("time"))
            if cvalue.lower().find("time") > -1:
                times.add(col)
        # 数据行
        rows = resultValue.get("rows", [])
        self.logger.debug("rows length: {0}".format(len(rows)))
        return times, rows

    def queryDb(self, dburl, query):
        # 连接数据库查询数据
        self.logger.debug("queryDb: query: %s" % query)
        engine = create_engine(dburl)
        query = query.replace("%", "%%")
        datas = engine.execute(query).fetchall()
        return datas
