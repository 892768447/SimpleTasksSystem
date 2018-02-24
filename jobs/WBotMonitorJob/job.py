#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月19日
@author: Irony."[讽刺]
@site: http://alyl.vip, http://orzorz.vip, https://coding.net/u/892768447, https://github.com/892768447
@email: 892768447@qq.com
@file: jobs.WBotMonitorJob.job
@description: 微信机器人
'''

import traceback

import AutoReportGlobals  # @UnresolvedImport
from AutoReportUtil import initLogger, getConfig  # @UnresolvedImport
from jobs.WBotMonitorJob.WBotMonitor import WBotMonitor  # @UnresolvedImport


__version__ = "0.0.1"

Enable = False  # 是否启用
Id = "WBotMonitorJob"  # 任务ID
Subject = "微信机器人"

def initjob():
    tname = "jobs/WBotMonitorJob/微信机器人配置文件.xlsx"  # 模版文件
    trigger, kwargs, config = getConfig(tname)  # @UnusedVariable

    # 初始日志文件
    logger = initLogger("datas/logs/" + Id, Id)
    wbmjob = WBotMonitor(logger, config, Id)  # 涉及到类全局变量,需要放在外面初始化

    @AutoReportGlobals.Scheduler.scheduled_job(trigger, id=Id, name=Subject, misfire_grace_time=86400, **kwargs)
    def job():
        try:
            wbmjob.start()
        except Exception as e:
            traceback.print_exc()
            logger.error("任务执行失败: " + str(e))

def init():
    if not Enable:
        return
    initjob()  # 初始任务
