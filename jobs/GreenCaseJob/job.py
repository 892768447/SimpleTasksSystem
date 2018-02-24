#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月10日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.GreenCaseJob
@description: 绿网案例入库采集
'''
from multiprocessing import Process
import traceback

from tornado.options import options

import AutoReportGlobals  # @UnresolvedImport
from AutoReportUtil import initLogger, getConfig  # @UnresolvedImport
from jobs.GreenCaseJob.GreenCaseMonitor import GreenCaseMonitor  # @UnresolvedImport
from jobs.GreenCaseJob.GreenCaseStorage import GreenCaseStorage  # @UnresolvedImport


__Author__ = "By: Irony.\"[讽刺]\nQQ: 892768447\nEmail: 892768447@qq.com"
__Copyright__ = "Copyright (c) 2017 Irony.\"[讽刺]"
__Version__ = "Version 1.0"

Enable = True  # 是否启用
Id = "GreenCaseJob"  # 任务名称
Subject = "绿网案例监控和入库"

def green_case_in(Id):  # 入库
    tname = "jobs/GreenCaseJob/绿网入库配置文件.xlsx"  # 模版文件
    trigger, kwargs, config = getConfig(tname)  # @UnusedVariable
    gcsjob = GreenCaseStorage(Id, config, tname)
    gcsjob.start()

def green_case_mi(Id):  # 监控
    tname = "jobs/GreenCaseJob/绿网监控配置文件.xlsx"  # 模版文件
    gcmjob = GreenCaseMonitor(Id, tname, options.port)
    gcmjob.start()

def initjob():

    Id1 = Id + "-1"
    tname1 = "jobs/GreenCaseJob/绿网入库配置文件.xlsx"  # 模版文件
    trigger1, kwargs1, _ = getConfig(tname1)  # @UnusedVariable

    @AutoReportGlobals.Scheduler.scheduled_job(trigger1, id=Id1, name="绿网案例入库", **kwargs1)
    def job_green_case_in():
        try:
            process = Process(target=globals()["green_case_in"], args=(Id1,))
            process.start()
            process.join()
        except SystemExit:pass
        except Exception as e:
            traceback.print_exc()
            logger = initLogger("datas/logs/" + Id1, Id1)
            logger.error("任务执行失败: " + str(e))

    Id2 = Id + "-2"
    tname2 = "jobs/GreenCaseJob/绿网监控配置文件.xlsx"  # 模版文件
    trigger2, kwargs2, _ = getConfig(tname2)

    @AutoReportGlobals.Scheduler.scheduled_job(trigger2, id=Id2, name="绿网案例监控", **kwargs2)
    def job_green_case_mi():
        try:
            process = Process(target=globals()["green_case_mi"], args=(Id2,))
            process.start()
            process.join()
        except SystemExit:pass
        except Exception as e:
            traceback.print_exc()
            logger = initLogger("datas/logs/" + Id2, Id2)
            logger.error("任务执行失败: " + str(e))


def init():
    if not Enable:
        return
    initjob()  # 初始任务
