#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年8月16日
@author: Irony."[讽刺]
@site: http://alyl.vip, http://orzorz.vip, https://coding.net/u/892768447, https://github.com/892768447
@email: 892768447@qq.com
@file: jobs.Reports.BroadbandDailyJob
@description: 岁末营销日报
'''
# import sys
# sys.path.insert(0, "../../../")
from datetime import datetime, timedelta
from multiprocessing import Process
import traceback

import AutoReportGlobals  # @UnresolvedImport
from AutoReportUtil import initLogger, getConfig  # @UnresolvedImport
from jobs.Reports.FileDaily import FileDaily  # @UnresolvedImport


__Author__ = "By: Irony.\"[讽刺]\nQQ: 892768447\nEmail: 892768447@qq.com"
__Copyright__ = "Copyright (c) 2017 Irony.\"[讽刺]"
__Version__ = "Version 1.0"

Enable = True  # 是否启用
Id = "FourthJob"  # 任务名称
Subject = "岁末营销日报"


def broad_band(Id, tname):
    trigger, kwargs, config = getConfig(tname)  # @UnusedVariable
    Subject = "岁末营销日报" + \
        (datetime.now() + timedelta(days=-2)).strftime(" - %Y%m%d")
    fdjob = FileDaily(Id, config, tname, Subject)
    fdjob.start()


def initjob():
    Id0 = Id + "-0"
    tname0 = "jobs/Reports/FourthJob/岁末营销日报.xlsx"
    trigger0, kwargs0, config0 = getConfig(tname0)  # @UnusedVariable

    @AutoReportGlobals.Scheduler.scheduled_job(trigger0, id=Id0, name=Subject + Id0, **kwargs0)
    def job_broad_band():
        try:
            process = Process(
                target=globals()["broad_band"], args=(Id0, tname0))
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


if __name__ == "__main__":
    tname = "岁末营销日报.xlsx"
    trigger, kwargs, config = getConfig(tname)
    initLogger("../../../datas/logs/" + Id, Id)
    fdjob = FileDaily(Id, config, tname, Subject)
    fdjob.start()
