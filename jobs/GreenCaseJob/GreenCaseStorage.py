#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月10日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.tasks.GreenCaseStorage
@description: 绿网案例入库
'''
from calendar import monthrange
from datetime import datetime, timedelta
import os
from random import randrange
import sys
from time import sleep
import traceback

import pymysql  # @UnusedImport
import requests
from sqlalchemy.engine import create_engine
from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.orm.scoping import scoped_session
from sqlalchemy.orm.session import sessionmaker
from sqlalchemy.pool import QueuePool
from sqlalchemy.sql.schema import Column
from sqlalchemy.sql.sqltypes import VARCHAR, TEXT
import xlrd

from AutoReportUtil import initLogger  # @UnresolvedImport
from AutoReportUtil import sendEmailText  # @UnresolvedImport
from jobs.GreenCaseJob.GreenCase import login  # @UnresolvedImport
from openpyxl.reader.excel import load_workbook


__version__ = "0.0.1"

DirPath = "datas/tmps/greencase/"
os.makedirs(DirPath, exist_ok=True)
# 表单
PARAM = "appealAcceptQuery.acceptId=&appealAcceptQuery.appealTel=&appealAcceptQuery.acceptTime=&appealAcceptQuery.startTime={startTime}&appealAcceptQuery.endTime={endTime}&appealAcceptQuery.acceptClassDescr=&appealAcceptQuery.acceptClass=&appealAcceptQuery.telArea=%C3%BC%C9%BD&appealAcceptQuery.status=&appealAcceptQuery.appealLevel=&appealAcceptQuery.ring=&appealAcceptQuery.acceptMode=&appealAcceptQuery.dealStaff=&deptDescr=&appealAcceptQuery.deptId=&appealAcceptQuery.staffId=&appealAcceptQuery.brand=&appealAcceptQuery.custmerType=&appealAcceptQuery.pourstates=&appealAcceptQuery.acceptSource=&appealAcceptQuery.duration=&appealAcceptQuery.isShowDetail=on&deptDescr2=&appealAcceptQuery.deptId2=&appealAcceptQuery.satisfied=0&appealAcceptQuery.acceptId=&appealAcceptQuery.appealTel="
# 投诉单子详情
DETAIL_URL = "http://10.109.214.54:8080/cscwf/wf/appealshowview/AppealShowAction.action?acceptId={acceptId}"
# 下载xls地址
XLS_URL = "http://10.109.214.54:8080/cscwf/wf/query/queryAppealAcceptExportAction.action"
HEADERS = {
    "Accept": "image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*",
    "Referer": "http://10.109.214.54:8080/cscwf/wf/query/gotoQueryAppealAcceptAction.action?module=M0006",
    "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)",
    "Content-Type": "application/x-www-form-urlencoded",
    "Cache-Control": "no-cache"
}
Base = declarative_base()


class Sheet1(Base):  # 数据库表结构

    __tablename__ = "sheet1"

    gdbh = Column("工单编号", VARCHAR(50), primary_key=True, autoincrement=False)
    tshm = Column("投诉号码", VARCHAR(15))
    slgh = Column("受理工号", VARCHAR(20))
    slsj = Column("受理时间", VARCHAR(50))
    tslx = Column("投诉类型", VARCHAR(255))
    tsnr = Column("投诉内容", TEXT)
    clgc = Column("处理过程", TEXT)
#     clxq = Column("处理详情", TEXT)
    gdyj = Column("归档意见", TEXT)
    hmgs = Column("号码归属地", VARCHAR(255))
    gzdd = Column("故障地点", VARCHAR(255))
    clzt = Column("处理状态", VARCHAR(20))
    gdbm = Column("归档部门", VARCHAR(50))
    gdgh = Column("归档工号", VARCHAR(20))
    gdbz = Column("归档工号所在班组", VARCHAR(50))
    gdsj = Column("归档时间", VARCHAR(50))
    myqk = Column("满意度情况", VARCHAR(100))
    cpcs = Column("不满意重派次数", VARCHAR(20))
    cpxq = Column("不满意重派详情", TEXT)

class GreenCaseStorage:
    
    def __init__(self, Id, config, tname):
        self.retrys = 0
        self.Id = Id
        self.logger = initLogger("datas/logs/" + Id, Id)
        self.config = config.get("配置", None)
        self.tname = tname
        wb = load_workbook(tname)
        # 预发送人
        preSheet = wb["预发送"]
        if not preSheet:
            self.logger.warn("无法找到预发送人表")
            self.preSend = []
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

    def getParam(self, startTime=None, endTime=None):
        return PARAM.format(startTime=startTime, endTime=endTime).encode()
    
    def parseConfig(self):
        self.logger.debug("parse config start")
        now = datetime.now()
        self.logger.info("**********%s**********" % 
                         now.strftime("%Y-%m-%d %H-%M-%S"))
        # 数据库地址
        databaseurl = self.config.get("数据库连接地址", None)  # ?charset=utf8
        databasename = self.config.get("数据库名", None)
        if (not databaseurl) or (not databasename):
            self.logger.error("没有找到数据库连接地址或者数据库名")
            return 0
        self.days = self.config.get("前几天", 5)
        self.logger.debug("前几天:%s" % self.days)
        self.timeout = self.config.get("超时时间", 120)
        self.logger.debug("超时时间:%s" % self.timeout)
        self.sleepnum = self.config.get("随机休眠时间", 5)
        self.logger.debug("随机休眠时间:%s" % self.sleepnum)
        self.threadnum = self.config.get("线程数", 5)
        self.logger.debug("线程数:%s" % self.threadnum)
        self.ispatch = self.config.get("补数据", None) == "是"
        self.logger.debug("补数据:%s" % self.ispatch)
        self.year = self.config.get("年", 0)
        self.logger.debug("年:%s" % self.year)
        self.month = self.config.get("月", 1)
        self.logger.debug("月:%s" % self.month)
        if self.ispatch and not self.year:
            self.logger.error("补数据的年份不能为空")
            return 0
        try:
            # 尝试创建数据库
            engine = create_engine(databaseurl)
            # 创建数据库
            with engine.connect() as con:
                con.execute(
                    "CREATE DATABASE IF NOT EXISTS {0} CHARACTER SET utf8;".format(
                        databasename))
            # 连接数据库
            engine = create_engine(
                databaseurl + databasename + "?charset=utf8",
                pool_size=self.threadnum,
                poolclass=QueuePool,
            )  # echo=True
            Base.metadata.create_all(engine)  # 创建表
            self.session = scoped_session(
                sessionmaker(bind=engine,
                             autocommit=False,
                             autoflush=True
                )
            )
        except Exception as e:
            traceback.print_exc()  # 初始化失败
            self.logger.error(str(e))
            return 0
        self.logger.debug("解析配置成功")
        return 1
    
    def start(self):
        if not self.config:
            return self.logger.error("无法获取配置文件")
        self.getLogin()
        # 登录成功
        if not self.parseConfig():  # 解析配置
            raise Exception("解析配置失败")
        # 下载数据
        self.logger.debug("downlaod start")
        if self.ispatch:  # 补数据
            for month in range(self.month, 13):
                startTime = "{year}-{month}-01+00%3A00%3A00".format(year=self.year,
                    month=str(month).zfill(2))
                endTime = "{year}-{month}-15+23%3A59%3A59".format(year=self.year,
                    month=str(month).zfill(2))
                print("startTime: ", startTime, " endTime: ", endTime)
                self._getData(startTime, endTime, False)
                sleep(randrange(self.sleepnum))
                startTime = "{year}-{month}-16+00%3A00%3A00".format(year=self.year,
                    month=str(month).zfill(2))
                endTime = "{year}-{month}-{day}+23%3A59%3A59".format(year=self.year,
                    month=str(month).zfill(2), day=monthrange(self.year, month)[1])
                print("startTime: ", startTime, " endTime: ", endTime)
                self._getData(startTime, endTime, False)
            sys.exit()
        else:
            # 目前这里为1提取前一天的数据
            now = datetime.now()
            if now.day < self.days:  # 每月前几天
                qnow = now + timedelta(days=-now.day + 1)
            else:
                qnow = now + timedelta(days=-self.days)
            StartTime = qnow.strftime("%Y-%m-%d") + "+00%3A00%3A00"
            EndTime = now.strftime("%Y-%m-%d+%H%%3A%M%%3A%S")
            
            self.logger.debug("StartTime:%s" % StartTime)
            self.logger.debug("StartTime:%s" % str(StartTime))
            self.logger.debug("EndTime:%s" % EndTime)
            self.logger.debug("EndTime:%s" % str(EndTime))
            self.logger.debug("定时采取:%s,%s" % (StartTime, EndTime))
            self._getData(StartTime, EndTime)
        
        # 发送邮件通知
        self.logger.debug(str(sendEmailText(self.preSend, "绿网入库", "任务:%s 已完成" % self.Id)))
    
    def _getData(self, startTime, endTime, ifquit=True):
        outFile = (DirPath + startTime + "--" + endTime + ".xls").replace("+", "-").replace("%3A", "-")
        self.logger.debug("outFile:%s" % outFile)
        req = requests.post(XLS_URL, data=self.getParam(startTime, endTime),
            headers=HEADERS, cookies=self.cookies, timeout=self.timeout)
        if req.text.find("版权所有") > 0:  # 需要重新登录
            self.logger.debug("需要重新登录")
            self.getLogin()
        with open(outFile, "wb") as fp:
            for chunk in req.iter_content(chunk_size=1024):
                if chunk: fp.write(chunk)
            fp.close()
        self.insertData(outFile, ifquit)

    def insertData(self, outFile, ifquit=True):
        book = xlrd.open_workbook(outFile)
        sheet = book.sheet_by_index(0)
        if not sheet.nrows:
            self.logger.debug("no data")
            return sys.exit() if ifquit else None
        self.logger.debug("行: %s" % str(sheet.nrows))
        self.logger.debug("insert start")
        session = self.session()
        for i in range(1, sheet.nrows):
            try:
                item = sheet.row_values(i)[:18]
                if not item[0]:
                    self.logger.debug("空值:%s" % str(i))
                    continue
                data = {
                    "gdbh":item[0], "tshm":item[1],
                    "slgh":item[2], "slsj":item[3],
                    "tslx":item[4], "tsnr":item[5],
                    "clgc":item[6], "gdyj":item[7],
                    "hmgs":item[8], "gzdd":item[9],
                    "clzt":item[10], "gdbm":item[11],
                    "gdgh":item[12], "gdbz":item[13],
                    "gdsj":item[14], "myqk":item[15],
                    "cpcs":item[16], "cpxq":item[17],
                }
                gdbh = item[0]
                if not gdbh: continue
                if not session.query(Sheet1).filter_by(gdbh=gdbh).all():  # 不存在插入
                    # print("插入新数据")
                    session.add(Sheet1(**data))
                else:  # 存在更新
                    # print("更新旧数据")
                    session.query(Sheet1).filter_by(gdbh=gdbh).update({
                        "tsnr":data.get("tsnr"),
                        "clgc":data.get("clgc")
                    })
            except Exception as e: 
                self.logger.error("row: %s error:%s" % (str(i), e))
        session.commit()
        session.close()
        self.logger.debug("insert end")
        if ifquit: sys.exit()
