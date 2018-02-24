#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月18日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.ReportCollectionJob.ReportCollection
@description: 
'''
'''
from CpBgCollection import CpBgCollection  # @UnresolvedImport
from CzMxCollection import CzMxCollection  # @UnresolvedImport
from YxBlCollection import YxBlCollection  # @UnresolvedImport
from YxCzCollection import YxCzCollection  # @UnresolvedImport
'''

from concurrent.futures.thread import ThreadPoolExecutor
from datetime import datetime
import os

import requests
import yaml

from AutoReportUtil import initLogger  # @UnresolvedImport


# import json
ConfigsDir = "jobs/ReportCollectionJob/configs"


__version__ = "0.0.1"


class ReportCollection:

    def __init__(self, Id, tname):
        self.logger = initLogger("datas/logs/" + Id, Id)

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
#         Codes = json.load(open("configs/codes.txt", "r", encoding="gbk"))
        Codes = yaml.load(
            open(os.path.join(ConfigsDir, "codes.txt"), "r", encoding="gbk").read())
        # 加载cookies文件
        Cookies = yaml.load(
            open(os.path.join(ConfigsDir, "cookies.txt"), "r", encoding="gbk").read())
#         Cookies = json.load(
#             open(os.path.join(ConfigsDir, "cookies.txt"), "r", encoding="gbk"))
        # 加载referers文件
        Referers = yaml.load(
            open(os.path.join(ConfigsDir, "referers.txt"), "r", encoding="gbk").read())
#         Referers = json.load(
#             open(os.path.join(ConfigsDir, "referers.txt"), "r", encoding="gbk"))
        # 加载网址
        Urls = yaml.load(
            open(os.path.join(ConfigsDir, "urls.txt"), "r", encoding="gbk").read())
#         Urls = json.load(
#             open(os.path.join(ConfigsDir, "urls.txt"), "r", encoding="gbk"))
        # 加载表单
        Params = yaml.load(
            open(os.path.join(ConfigsDir, "params.txt"), "r", encoding="gbk").read())
#         Params = json.load(
#             open(os.path.join(ConfigsDir, "params.txt"), "r", encoding="gbk"))
        # 加载配置
        Config = yaml.load(
            open(os.path.join(ConfigsDir, "config.txt"), "r", encoding="gbk").read())
#         Config = json.load(
# open(os.path.join(ConfigsDir, "config.txt"), "r", encoding="gbk"))

        # 随机数
        RandomNum = str(datetime.now().day).zfill(2) + "090230123"

        # 第一步先获取动态的host
        self.form_bg = "http://{host}/sireports/userdefined_reports/css/ng/nresources/UI/images/form_bg.png"
        self.hosts = Urls.get("hosts", "").split(";")
#         print("hosts: ", hosts)
        self.logger.debug("hosts: %s" % self.hosts)

        # 动态加载采集模块
        modules = Config.get("modules", {})
#         print("modules:", modules)
        self.logger.debug("modules: %s" % modules)
        self.ModulesCollection = []
        for key, value in modules.items():
            self.logger.info("load module: %s %s" % (key, value))
#             print("load module: ", key, value, "modules." + key)
            # __import__(key, fromlist=("modules", key))
            Class = __import__(
                "jobs.ReportCollectionJob.modules." + key, fromlist=[key])
#             print("Class: ", Class)
            self.logger.debug("Class: %s" % Class)
            self.ModulesCollection.append(getattr(Class, key)(self.logger, Config.copy(
            ), Codes, Headers.copy(), Cookies, Referers, RandomNum, Urls, Params, tname))

#         self.yxbl = YxBlCollection(logger, Config.copy(), Codes, Headers.copy(), Cookies, Referers, RandomNum, Urls, Params)
#         self.yxcz = YxCzCollection(logger, Config.copy(), Codes, Headers.copy(), Cookies, Referers, RandomNum, Urls, Params)
#         self.czmx = CzMxCollection(logger, Config.copy(), Codes, Headers.copy(), Cookies, Referers, RandomNum, Urls, Params)
#         self.cpbg = CpBgCollection(logger, Config.copy(), Codes, Headers.copy(), Cookies, Referers, RandomNum, Urls, Params)

        # 多线程执行器
        self.executor = ThreadPoolExecutor(max_workers=5)

    def start(self):
        self.logger.info("任务开始")
        self.logger.debug("开始检测可用抓取地址")
        Host = None
        for host in self.hosts:
            try:  # 动态检测
                req = requests.get(self.form_bg.format(host=host), timeout=20)
                if req.status_code == 200:
                    Host = host
                    break
            except:
                pass
#         print("use host: ", Host)
        self.logger.info("use host: %s" % Host)
        if not Host:
            return self.logger.warn("无法找到可用的抓取地址")
        for Collection in self.ModulesCollection:
            #             print(Collection)
            self.executor.submit(Collection.run, Host)
#         self.executor.submit(self.yxbl.run)  # 营销业务所有办理报表
#         self.executor.submit(self.yxcz.run)  # 营销业务所有冲正报表
#         self.executor.submit(self.czmx.run)  # 营业员操作明细报表
#         self.executor.submit(self.cpbg.run)  # 产品变更明细统计
        self.executor.shutdown(wait=True)  # 等待所有任务完成
        self.logger.info("本次采集任务已完成")
