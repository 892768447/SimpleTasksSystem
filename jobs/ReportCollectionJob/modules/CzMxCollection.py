#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年6月14日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.ReportCollectionJob.modules.YxBlCollection
@description: 营销业务所有操作明细报表
'''
import sys
sys.path.insert(0, "../../../")

from datetime import datetime

from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.sql.schema import Column
from sqlalchemy.sql.sqltypes import Integer, String

from jobs.ReportCollectionJob.BaseCollection import BaseCollection  # @UnresolvedImport


__version__ = "0.0.1"
# 创建对象的基类:
Base = declarative_base()

# 定义数据库模型


class CzMxCollectionModel(Base):

    @classmethod
    def keys(self):
        return ["fwhm", "qx", "yyt", "czmc", "czgh", "czls",
                "czsj", "ywpp", "sxf", "skf", "qtf", "jehj"
                ]
    # 库名
    __dbname__ = "BossReport".lower()
    # 表名 yxbl_201706
    __tablename__ = "CzMx".lower()

    # 表结构
    id = Column(Integer(), primary_key=True)
    fwhm = Column(String(100), doc="服务号码")
    qx = Column(String(100), doc="区县")
    yyt = Column(String(100), doc="营业厅")
    czmc = Column(String(100), doc="操作名称")
    czgh = Column(String(100), doc="操作工号")
    czls = Column(String(100), doc="操作流水")
    czsj = Column(String(100), doc="操作时间")
    ywpp = Column(String(100), doc="业务品牌")
    sxf = Column(String(100), doc="手续费")
    skf = Column(String(100), doc="SIM卡费")
    qtf = Column(String(100), doc="其他费")
    jehj = Column(String(100), doc="金额合计")


class CzMxCollection(BaseCollection):

    Name = "营业员操作明细报表"  # 名字
    TmpDir = "datas/tmps/营业员操作明细报表"  # tmp目录
    DataDir = "datas/营业员操作明细报表"  # data目录

    def getModel(self):
        # 数据库模型
        return CzMxCollectionModel

    def getBase(self):
        return Base

    def getTime(self):
        # 删除条件的时间字段
        return CzMxCollectionModel.czsj


if __name__ == "__main__":
    import logging
    import yaml
    import requests
    handler = logging.FileHandler(
        "CzMxCollection.log", mode="w", encoding="gbk")
    formatter = logging.Formatter(
        "[%(asctime)s %(name)s:%(lineno)s] P:%(process)d T:%(thread)d %(levelname)-8s %(message)s<br />")
    handler.setFormatter(formatter)
    logger = logging.getLogger("CzMxCollection")
    logger.addHandler(handler)
    logger.setLevel(logging.DEBUG)

    # 加载模拟头
    Headers = {
        "Accept": "image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*",
        "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; InfoPath.2)",
        "Content-Type": "application/x-www-form-urlencoded",
        "Connection": "Keep-Alive",
        "Cache-Control": "no-cache",
        "Referer": "",
        "Cookie": "",
    }
    # 加载地区指标代码
    Codes = yaml.load(open("../configs/codes.txt", "r", encoding="gbk").read())
    # 加载cookies文件
    Cookies = yaml.load(open("../configs/cookies.txt",
                             "r", encoding="gbk").read())
    # 加载referers文件
    Referers = yaml.load(
        open("../configs/referers.txt", "r", encoding="gbk").read())
    # 加载网址
    Urls = yaml.load(open("../configs/urls.txt", "r", encoding="gbk").read())
    # 加载表单
    Params = yaml.load(open("../configs/params.txt",
                            "r", encoding="gbk").read())
    # 加载配置
    Config = yaml.load(open("../configs/config.txt",
                            "r", encoding="gbk").read())
    # 随机数
    RandomNum = str(datetime.now().day).zfill(2) + "090230123"  # 105617424
    # 第一步先获取动态的host
    form_bg = "http://{host}/sireports/userdefined_reports/css/ng/nresources/UI/images/form_bg.png"
    hosts = Urls.get("hosts", "").split(";")
    logger.debug("hosts: %s" % hosts)
    for host in hosts:
        try:  # 动态检测
            req = requests.get(form_bg.format(host=host), timeout=20)
            if req.status_code == 200:
                Host = host
                break
        except:
            pass
    logger.info("use host: %s" % Host)
    tname = "../营销报表提取配置文件1200.xlsx"
    job = CzMxCollection(logger, Config.copy(), Codes, Headers.copy(
    ), Cookies, Referers, RandomNum, Urls, Params, tname)
    job.run(Host)
    print("end")
