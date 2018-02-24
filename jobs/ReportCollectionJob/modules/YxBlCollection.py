#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年6月14日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.ReportCollectionJob.modules.YxBlCollection
@description: 营销业务所有办理报表
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


class YxBlCollectionModel(Base):

    @classmethod
    def keys(self):
        return ["gh", "qx", "yxzx", "yyt", "qdbm",
                "qdlx", "zdjx", "zdch", "blsj", "fsbm",
                "zdlx", "zdzs", "hdmc", "fsmc", "ywlx",
                "ddh", "jfje", "syzkqje", "gjk", "yck",
                "zsyhje", "zsycje", "zskje", "zswje",
                "zsqje", "zsdkqje", "sfydd", "fwhm",
                "ywdm", "jffs", "zdxh", "cgj", "dxsthsj",
                "mdj", "yhsxxxsj", "dbyf", "hylx", "hyj",
                "bhsgjk", "sj", "sfkqbl"
                ]
    # 库名
    __dbname__ = "BossReport".lower()
    # 表名 yxbl_201706
    __tablename__ = "YxBl".lower()

    # 表结构
    id = Column(Integer(), primary_key=True)
    gh = Column(String(100), doc="工号")
    qx = Column(String(100), doc="区县")
    yxzx = Column(String(100), doc="营销中心")
    yyt = Column(String(100), doc="营业厅")
    qdbm = Column(String(100), doc="渠道编码")
    qdlx = Column(String(100), doc="渠道类型")
    zdjx = Column(String(100), doc="终端机型")
    zdch = Column(String(100), doc="终端串号")
    blsj = Column(String(100), doc="办理时间")
    fsbm = Column(String(100), doc="方式编码")
    zdlx = Column(String(100), doc="终端类型")
    zdzs = Column(String(100), doc="终端制式")
    hdmc = Column(String(100), doc="活动名称")
    fsmc = Column(String(100), doc="方式名称")
    ywlx = Column(String(100), doc="业务类型")
    ddh = Column(String(100), doc="订单号")
    jfje = Column(String(100), doc="缴费金额")
    syzkqje = Column(String(100), doc="使用抵扣券金额")
    gjk = Column(String(100), doc="购机款")
    yck = Column(String(100), doc="预存款")
    zsyhje = Column(String(100), doc="赠送优惠金额")
    zsycje = Column(String(100), doc="赠送预存金额")
    zskje = Column(String(100), doc="赠送卡金额")
    zswje = Column(String(100), doc="赠送物金额")
    zsqje = Column(String(100), doc="赠送其金额")
    zsdkqje = Column(String(100), doc="赠送抵扣券金额")
    sfydd = Column(String(100), doc="是否有订单")
    fwhm = Column(String(100), doc="服务号码")
    ywdm = Column(String(100), doc="业务代码")
    jffs = Column(String(100), doc="缴费方式")
    zdxh = Column(String(100), doc="终端型号")
    cgj = Column(String(100), doc="采购价")
    dxsthsj = Column(String(100), doc="代销商提货时间")
    mdj = Column(String(100), doc="买断价")
    yhsxxxsj = Column(String(100), doc="用户实现销售时间")
    dbyf = Column(String(100), doc="打标与否")
    hylx = Column(String(100), doc="合约类型")
    hyj = Column(String(100), doc="合约价")
    bhsgjk = Column(String(100), doc="不含税购机款")
    sj = Column(String(100), doc="税金")
    sfkqbl = Column(String(100), doc="是否跨区办理")


class YxBlCollection(BaseCollection):

    Name = "营销业务所有办理报表"  # 名字
    TmpDir = "datas/tmps/营销业务所有办理报表"  # tmp目录
    DataDir = "datas/营销业务所有办理报表"  # data目录

    def getModel(self):
        # 数据库模型
        return YxBlCollectionModel

    def getBase(self):
        return Base

    def getTime(self):
        # 删除条件的时间字段
        return YxBlCollectionModel.blsj


if __name__ == "__main__":
    import logging
    import yaml
    import requests
    handler = logging.FileHandler(
        "YxBlCollection.log", mode="w", encoding="gbk")
    formatter = logging.Formatter(
        "[%(asctime)s %(name)s:%(lineno)s] P:%(process)d T:%(thread)d %(levelname)-8s %(message)s<br />")
    handler.setFormatter(formatter)
    logger = logging.getLogger("YxBlCollection")
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

    SPLIT = Config.get("split", 0)
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
    job = YxBlCollection(logger, Config.copy(), Codes, Headers.copy(
    ), Cookies, Referers, RandomNum, Urls, Params)
    job.run(Host)
    print("end")
