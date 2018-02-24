#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月4日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.EmailMonitorJob
@description: 邮件监听-负责命令控制
'''
import traceback

import AutoReportGlobals  # @UnresolvedImport
from AutoReportUtil import initLogger, getConfig  # @UnresolvedImport
from jobs.EmailMonitorJob.EmailMonitor import EmailMonitor  # @UnresolvedImport


__version__ = "0.0.1"

Enable = True  # 是否启用
Id = "EmailMonitorJob"  # 任务ID
Subject = "邮件服务"


def initjob():
    tname = "jobs/EmailMonitorJob/邮件监控配置文件.xlsx"  # 模版文件
    trigger, kwargs, config = getConfig(tname)  # @UnusedVariable

    # 初始日志文件
    logger = initLogger("datas/logs/" + Id, Id)
    emailjob = EmailMonitor(logger)  # 涉及到类全局变量,需要放在外面初始化

    # minutes -- 每1分钟监听一次
    # seconds -- 每40秒监听一次
    # @AutoReportGlobals.Scheduler.scheduled_job("interval", id=Id, name="邮件监听", minutes=1)
    # 修改为每天早上8点开始到下午19点每隔1分钟
    # @AutoReportGlobals.Scheduler.scheduled_job("cron", id=Id, name="邮件监听", minute="0-59", hour="8-19")
    @AutoReportGlobals.Scheduler.scheduled_job(trigger, id=Id, name=Subject, **kwargs)
    def job_email_monitor():
        try:
            emailjob.start()
        except Exception as e:
            traceback.print_exc()
            logger.error("任务执行失败: " + str(e))


def init():
    if not Enable:
        return
    initjob()  # 初始任务
