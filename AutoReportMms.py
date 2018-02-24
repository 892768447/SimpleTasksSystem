#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年6月29日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: AutoReportMms
@description: 
'''
import logging
import os
import traceback

from apscheduler.events import EVENT_JOB_EXECUTED, EVENT_JOB_ERROR, \
    EVENT_JOB_MISSED
from apscheduler.executors.pool import ThreadPoolExecutor  # , ProcessPoolExecutor
from apscheduler.schedulers.tornado import TornadoScheduler
import requests
from tornado.httpserver import HTTPServer
from tornado.ioloop import IOLoop
from tornado.options import define, options

from AutoReportApplication import AutoReportApplication  # @UnresolvedImport
import AutoReportGlobals  # @UnresolvedImport
from AutoReportUtil import initJobs  # @UnresolvedImport


# from apscheduler.jobstores.sqlalchemy import SQLAlchemyJobStore
__version__ = "0.0.1"

define("port", default=8989,
       help="run on the given port,default is 8989", type=int)
define("address", default="0.0.0.0", help="run on the given address", type=str)
define("stop", default=False, help="stop the server", type=bool)

DbStoreTest = "sqlite:///datas/data.sqlite3"
DbStore = "mysql+pymysql://root:root@localhost/taskstore"

os.makedirs("datas/tmps", exist_ok=True)  # 临时目录
os.makedirs("datas/logs", exist_ok=True)  # 日志目录
os.makedirs("datas/web/template", exist_ok=True)  # web模版路径
os.makedirs("datas/web/static", exist_ok=True)  # web静态资源路径

# 日志文件
# handler = logging.FileHandler("datas/logs/AutoReportMms" + 
#     datetime.now().strftime("%Y-%m-%d") + ".log", mode="w", encoding="gbk")
handler = logging.FileHandler("datas/logs/AutoReportMms.log", mode="w", encoding="gbk")
formatter = logging.Formatter(
    "[%(asctime)s %(filename)s:%(lineno)s] P:%(process)d T:%(thread)d %(levelname)-8s %(message)s<br />")
handler.setFormatter(formatter)

# 修改aps的logger
loggeraps = logging.getLogger("apscheduler.executors.default")
loggeraps.addHandler(handler)
loggeraps.setLevel(logging.DEBUG)

# 自己的logger
logger = logging.getLogger("AutoReportMms")
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)

# 采用数据库存储
# jobstores = {
#     "default": SQLAlchemyJobStore(DbStore)
# }
executors = {
    "default": ThreadPoolExecutor(10),  # 线程池
#     "processpool": ProcessPoolExecutor(5)  # 进程池
}
job_defaults = {
    "coalesce": True,  # 只执行一次
    "max_instances": 1  # 最大实例数量
}

# 实例化任务调度器TornadoScheduler
scheduler = TornadoScheduler(
    daemonic=False,  # 不使用守护进程
    #     jobstores=jobstores, #默认使用内存
    executors=executors, job_defaults=job_defaults)

# 全局可用
AutoReportGlobals.Scheduler = scheduler
AutoReportGlobals.EmailIds = {}


def onListener(event):
    if event.exception:
        logger.warn("The job crashed :( " + str(event))
#         print("The job crashed :(", event.alias, event.code, event.exception,
# event.job_id, event.jobstore, event.retval, event.scheduled_run_time)


def stop():
    try:
        scheduler.shutdown(wait=False)
    except Exception as e:
        logger.warn(str(e))
    IOLoop.instance().stop()


def start():
    # 解析命令行
    options.parse_command_line()
    if options.stop:
        return stop()
    
    # 先检测程序是否在运行
    try:
#         print("".join(("http://", options.address, ":", str(options.port), "/live")))
        u = "".join(("http://127.0.0.1:", str(options.port), "/live"))
        req = requests.get(u, timeout=10)
        if req.status_code == 200:
            return logger.warn("调度器程序已在运行")
    except Exception as e:
        logger.error(str(e))

    # 加载任务
    initJobs(logger)
    # 添加事件监听
    scheduler.add_listener(
        onListener, EVENT_JOB_EXECUTED | EVENT_JOB_ERROR | EVENT_JOB_MISSED)
    logger.info("scheduler started")
    # 启动调度器
    scheduler.start()
    logger.info("scheduler running")

    # web 监控服务
    logger.info("http server init")
    server = HTTPServer(AutoReportApplication(logger, scheduler))
    server.listen(options.port, options.address)
    logger.info(
        "run http server {0}:{1}".format(options.address, options.port))

    # 启动事件循环
    logger.info("io loop start")
    IOLoop.instance().start()

if __name__ == "__main__":
    try:
        start()
    except (KeyboardInterrupt, SystemExit) as e:
        logger.warn(str(e))
    except Exception as e:
        logger.error(str(e))
        traceback.print_exc()
