#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年6月14日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.ReportCollectionJob.modules.CpBgCollection
@description: 产品变更明细统计
'''
from datetime import datetime
import sys
sys.path.insert(0, "../../../")
from sqlalchemy.ext.declarative.api import declarative_base
from sqlalchemy.sql.schema import Column
from sqlalchemy.sql.sqltypes import Integer, String

from jobs.ReportCollectionJob.BaseCollection import BaseCollection  # @UnresolvedImport

__version__ = "0.0.1"

# 创建对象的基类:
Base = declarative_base()

# 定义数据库模型


class CpBgCollectionModel(Base):

    @classmethod
    def keys(self):
        return ["gs", "yhhm", "yhpp", "ywfs", "tcbs",
                "tcmc", "czy", "czsj", "czlx"
                ]
    # 库名(一般不修改)
    __dbname__ = "BossReport".lower()
    # 表名 cpbg_201706
    __tablename__ = "CpBg".lower()

    # 表结构
    id = Column(Integer(), primary_key=True)
    gs = Column(String(100), doc="归属")
    yhhm = Column(String(100), doc="用户号码")
    yhpp = Column(String(100), doc="用户品牌")
    ywfs = Column(String(100), doc="业务方式")
    tcbs = Column(String(100), doc="套餐标识")
    tcmc = Column(String(100), doc="套餐名称")
    czy = Column(String(100), doc="操作员")
    czsj = Column(String(100), doc="操作日期")
    czlx = Column(String(100), doc="操作类型")


class CpBgCollection(BaseCollection):

    Name = "产品变更明细统计"  # 名字
    TmpDir = "datas/tmps/产品变更明细统计"  # tmp目录
    DataDir = "datas/产品变更明细统计"  # data目录

    def getModel(self):
        # 数据库模型
        return CpBgCollectionModel

    def getBase(self):
        return Base

    def getTime(self):
        # 删除条件的时间字段
        return CpBgCollectionModel.czsj


if __name__ == "__main__":
    import logging
    import yaml
    import requests
    handler = logging.FileHandler(
        "CpBgCollection.log", mode="w", encoding="gbk")
    formatter = logging.Formatter(
        "[%(asctime)s %(name)s:%(lineno)s] P:%(process)d T:%(thread)d %(levelname)-8s %(message)s<br />")
    handler.setFormatter(formatter)
    logger = logging.getLogger("CpBgCollection")
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
    RandomNum = str(datetime.now().day).zfill(2) + "090230123"
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
    tname = "../营销报表提取配置文件每个小时.xlsx"
    job = CpBgCollection(logger, Config.copy(), Codes, Headers.copy(
    ), Cookies, Referers, RandomNum, Urls, Params, tname)
    job.run(Host)
    print("end")
