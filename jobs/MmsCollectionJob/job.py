#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月17日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.MmsCollectionJob
@description: 经分数据抓取任务
'''
from multiprocessing import Process
import traceback

import AutoReportGlobals  # @UnresolvedImport
from AutoReportUtil import initLogger, getConfig  # @UnresolvedImport
from jobs.MmsCollectionJob.MmsCollection import MmsCollection  # @UnresolvedImport


__Author__ = "By: Irony.\"[讽刺]\nQQ: 892768447\nEmail: 892768447@qq.com"
__Copyright__ = "Copyright (c) 2017 Irony.\"[讽刺]"
__Version__ = "Version 1.0"

Enable = True  # 是否启用
Id = "MmsCollectionJob"  # 任务名称
Subject = "经分提取"


def mms_collection(Id, tname):
    mcjob = MmsCollection(Id)
    mcjob.start(tname)


def initjob():

    Id0 = Id + "0830"
    tname0 = "jobs/MmsCollectionJob/经分提取配置文件0830.xlsx"  # 模版文件
    trigger0, kwargs0, _ = getConfig(tname0)
    # 初始日志文件

    @AutoReportGlobals.Scheduler.scheduled_job(trigger0, id=Id0, name=Subject + Id0, **kwargs0)
    def job_mms_collection0():
        try:
            process = Process(
                target=globals()["mms_collection"], args=(Id0, tname0))
            process.start()
            process.join()
        except Exception as e:
            traceback.print_exc()
            logger = initLogger("datas/logs/" + Id0, Id0)
            logger.error("任务执行失败: " + str(e))

    Id1 = Id + "1320"
    tname1 = "jobs/MmsCollectionJob/经分提取配置文件1320.xlsx"  # 模版文件
    trigger1, kwargs1, _ = getConfig(tname1)
    # 初始日志文件

    @AutoReportGlobals.Scheduler.scheduled_job(trigger1, id=Id1, name=Subject + Id1, **kwargs1)
    def job_mms_collection1():
        try:
            process = Process(
                target=globals()["mms_collection"], args=(Id1, tname1))
            process.start()
            process.join()
        except Exception as e:
            traceback.print_exc()
            logger = initLogger("datas/logs/" + Id1, Id1)
            logger.error("任务执行失败: " + str(e))

    Id2 = Id + "1620"
    tname2 = "jobs/MmsCollectionJob/经分提取配置文件1620.xlsx"  # 模版文件
    trigger2, kwargs2, _ = getConfig(tname2)
    # 初始日志文件

    @AutoReportGlobals.Scheduler.scheduled_job(trigger2, id=Id2, name=Subject + Id2, **kwargs2)
    def job_mms_collection2():
        try:
            process = Process(
                target=globals()["mms_collection"], args=(Id2, tname2))
            process.start()
            process.join()
        except Exception as e:
            traceback.print_exc()
            logger = initLogger("datas/logs/" + Id2, Id2)
            logger.error("任务执行失败: " + str(e))

    Id3 = Id + "1830"
    tname3 = "jobs/MmsCollectionJob/经分提取配置文件1830.xlsx"  # 模版文件
    trigger3, kwargs3, _ = getConfig(tname3)
    # 初始日志文件

    @AutoReportGlobals.Scheduler.scheduled_job(trigger3, id=Id3, name=Subject + Id3, **kwargs3)
    def job_mms_collection3():
        try:
            process = Process(
                target=globals()["mms_collection"], args=(Id3, tname3))
            process.start()
            process.join()
        except Exception as e:
            traceback.print_exc()
            logger = initLogger("datas/logs/" + Id3, Id3)
            logger.error("任务执行失败: " + str(e))

    Id4 = Id + "xiagnzhen"
    tname4 = "jobs/MmsCollectionJob/乡镇经分提取配置文件.xlsx"  # 模版文件
    trigger4, kwargs4, _ = getConfig(tname4)
    # 初始日志文件

    @AutoReportGlobals.Scheduler.scheduled_job(trigger4, id=Id4, name=Subject + Id4, **kwargs4)
    def job_mms_collection4():
        try:
            process = Process(
                target=globals()["mms_collection"], args=(Id4, tname4))
            process.start()
            process.join()
        except Exception as e:
            traceback.print_exc()
            logger = initLogger("datas/logs/" + Id4, Id4)
            logger.error("任务执行失败: " + str(e))


def init():
    if not Enable:
        return
    initjob()  # 初始任务
