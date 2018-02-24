#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月10日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.ReportCollectionJob
@description: 每日BOSS报表提取
'''
from multiprocessing import Process
import traceback

import AutoReportGlobals  # @UnresolvedImport
from AutoReportUtil import initLogger, getConfig  # @UnresolvedImport
from jobs.ReportCollectionJob.ReportCollection import ReportCollection  # @UnresolvedImport


__Author__ = "By: Irony.\"[讽刺]\nQQ: 892768447\nEmail: 892768447@qq.com"
__Copyright__ = "Copyright (c) 2017 Irony.\"[讽刺]"
__Version__ = "Version 1.0"

Enable = True  # 是否启用
Id = "ReportCollectionJob"  # 任务名称
Subject = "营销报表提取"


def report_collection(Id, tname):
    rcjob = ReportCollection(Id, tname)
    rcjob.start()


def initjob():
    Id0 = Id + "-0"
    tname0 = "jobs/ReportCollectionJob/营销报表提取配置文件每个小时.xlsx"
    trigger0, kwargs0, config0 = getConfig(tname0)  # @UnusedVariable

    @AutoReportGlobals.Scheduler.scheduled_job(trigger0, id=Id0, name=Subject + Id0, **kwargs0)
    def job_report_collection0():
        try:
            process = Process(
                target=globals()["report_collection"], args=(Id0, tname0))
            process.start()
            process.join()
        except Exception as e:
            traceback.print_exc()
            logger = initLogger("datas/logs/" + Id0, Id0)
            logger.error("任务执行失败: " + str(e))


def init():
    if not Enable:
        return
    initjob()  # 初始任务
