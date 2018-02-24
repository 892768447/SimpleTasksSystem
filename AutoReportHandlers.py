#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月1日
@author: Irony."[讽刺]
@site: http://alyl.vip, http://orzorz.vip, https://coding.net/u/892768447, https://github.com/892768447
@email: 892768447@qq.com
@file: AutoReportHandlers
@description: 
'''
from datetime import datetime
import json
import logging
import os
from time import time

from tornado.ioloop import IOLoop
from tornado.web import RequestHandler, asynchronous, StaticFileHandler
from tornado.websocket import WebSocketHandler

import AutoReportGlobals  # @UnresolvedImport
from AutoReportUtil import initJobs  # @UnresolvedImport


__Author__ = "By: Irony.\"[讽刺]\nQQ: 892768447\nEmail: 892768447@qq.com"
__Copyright__ = "Copyright (c) 2017 Irony.\"[讽刺]"
__Version__ = "Version 1.0"


class BaseHandler(RequestHandler):

    def get(self, *args, **kwargs):
        self.redirect("/")

    def set_default_headers(self):
        self.set_header("Server", "AutoReportMms Server QQ:892768447")


class IndexHandler(BaseHandler):

    def get(self, *args, **kwargs):
        scheduler = self.application.scheduler
        self.render("index.html", jobs=scheduler.get_jobs())

    @asynchronous
    def post(self, *args, **kwargs):
        try:
            jobs = [{
                "id": str(job.id),
                "args": str(job.args),
                "name": str(job.name),
                "trigger": str(job.trigger),
                "next_run_time": str(job.next_run_time),
            } for job in self.application.scheduler.get_jobs()]
            return self.finish({"code": 200, "error": "", "data": jobs})
        except Exception as e:
            self.finish({"code": 200, "error": e, "data": []})

class LiveHandler(RequestHandler):
    
    def get(self, *args, **kwargs):
        self.finish("live")

class ModifyHandler(BaseHandler):
    
    def get(self, *args, **kwargs):
        self.finish("暂未开发")

    @asynchronous
    def post(self, *args, **kwargs):
        pass


class ControlHandler(BaseHandler):

    @asynchronous
    def post(self, *args, **kwargs):
        jobid = self.get_argument("jobid", None)
        ctype = self.get_argument("type", None)
        if jobid is None or ctype is None:
            return self.finish({"code": 403, "msg": "jobid or type is required"})
        try:
            ctype = int(ctype)
        except:
            return self.finish({"code": 403, "msg": "type must be number"})
        try:
            job = self.application.scheduler.get_job(jobid)
            if ctype == 1:  # 查看
#                 jname = job.id.split("-")[0]
#                 filename = "/logs/{0}/{1}-{2}.log".format(
#                     job.id, job.id, datetime.now().strftime("%Y-%m-%d"))
                filename = "/logs/{0}.log".format(job.id)
                return self.finish({"code": 200, "msg": filename})
# return self.finish({"code": 200, "msg": str(job.__getstate__())})
            if ctype == 2:  # pause
                job.pause()
                return self.finish({"code": 200, "msg": "任务暂停成功"})
            if ctype == 3:  # resume
                job.resume()
                return self.finish({"code": 200, "msg": "任务恢复成功"})
            if ctype == 4:  # remove
                job.remove()
                return self.finish({"code": 200, "msg": "任务删除成功"})
            if ctype == 5:  # restart
                pass
        except Exception as e:
            return self.finish({"code": 403, "msg": "get job error: " + str(e)})
        self.finish({"code": 200, "msg": "未知状况"})


class RestartHandler(BaseHandler):

    def get(self, *args, **kwargs):
        try:
            # 移除所有的任务
            self.application.scheduler.remove_all_jobs()
            # 重新加载任务
            initJobs(logging.getLogger("AutoReportMms"))
            self.finish({"code": 200, "msg": "重启调度器成功"})
        except Exception as e:
            self.finish({"code": 403, "msg": "重启调度器失败: " + str(e)})

class LogsHandler(StaticFileHandler):

    def set_default_headers(self):
        self.set_header("Server", "AutoReportMms Server QQ:892768447")

    def set_header(self, name, value):
        if name == "Content-Type":
            value = "text/html; charset=gb2312"
        super(LogsHandler, self).set_header(name, value)


# class LogsHandler(StaticFileHandler):
#     
#     def set_header(self, name, value):
#         if name == "Content-Type":
#             value = "text/html; charset=UTF-8"
#         super(LogsHandler, self).set_header(name, value)

#     def get(self, path=None):
#         if not path:
#             return self.finish("文件未找到")
#         self.render("log.html", title=os.path.basename(path), path=path)

class WeChatHandler(BaseHandler):

    # 调用微信机器人发送消息
    def post(self, *args, **kwargs):
#         print(AutoReportGlobals.WeixinBot)
        if not AutoReportGlobals.WeixinBot:
            return self.finish({"code":-1, "msg":"机器人死了"})
        data = json.loads(self.request.body.decode())
        if not data:return self.finish({"code":-1, "msg":"没有提交数据"})
        try:
            typ = data.get("type", "")  # 类型
            data = data.get("data", "")  # 内容
            if typ == "text":  # 发送文字
                code = AutoReportGlobals.WeixinBot.sendText(data)
                return self.finish({"code":code, "msg":"发送文本消息完成"})
            if typ == "image":  # 发送图片
                if not isinstance(data, list):
                    return self.finish({"code":-1, "msg":"图片列表参数错误"})
                code = AutoReportGlobals.WeixinBot.sendImages(data)
                return self.finish({"code":code, "msg":"发送图片消息完成"})
        except Exception as e:
            self.finish({"code":-1, "msg":str(e)})

class EmailIdHandler(BaseHandler):
    
    def post(self, *args, **kwargs):
        data = json.loads(self.request.body.decode())
        self.application.logger.debug("data: %s" % data)
#         print(data)
#         print(os.getpid(), AutoReportGlobals)
#         print(id(AutoReportGlobals.EmailIds))
        typ = data.get("type", 1)
        eid = data.get("eid", None)
        einfo = data.get("einfo", None)
#         print(typ, eid, einfo)
        self.application.logger.debug("typ: %s, eid: %s, einfo: %s" % (typ, eid, einfo))
        try: typ = int(typ)
        except: return self.finish({"code":-1, "msg":"type类型错误"})
        if typ == 0:  # 储存id
            if not eid or not einfo:
                print({"code":-1, "msg":"未找到eid或einfo"})
                return self.finish({"code":-1, "msg":"未找到eid或einfo"})
            AutoReportGlobals.EmailIds[eid] = einfo
            self.application.logger.debug("缓存id成功:%s" % eid)
#             print({"code":1, "msg":"保存成功"}, AutoReportGlobals.EmailIds[eid])
            return self.finish({"code":1, "msg":"保存成功"})
        if typ == 1:  # 查询id
            return self.finish({"code":1, "einfo":
                AutoReportGlobals.EmailIds.get("eid", {})})
#         print({"code":-1, "msg":"some thing may be wrong"})
        self.finish({"code":-1, "msg":"some thing may be wrong"})

class QrcodeHandler(BaseHandler):
    
    ResultJson = {
        "status": 1,
        "msg": "",
        "title": "微信机器人登录二维码",
        "id": 1,
        "start": 0,
        "data": [
            {
                "alt": "扫一扫登录微信",
                "pid": 1,
                "src": "/static/qr.png",
                "thumb": ""
            }
        ]
    }
    
    def get(self, *args, **kwargs):
        if not os.path.isfile("datas/web/static/qr.png"):
            self.ResultJson["status"] = 0
            self.ResultJson["data"] = []
            return self.finish(self.ResultJson)
        self.finish(self.ResultJson)

class LogMmsHandler(BaseHandler):
    
    @asynchronous
    def get(self, *args, **kwargs):
        # 当天日志
        file = datetime.now().strftime(AutoReportGlobals.NameNowSMS)
        path = os.path.join(AutoReportGlobals.PathNowSMS, file)
        if not os.path.isfile(path):
            return self.finish("日志未找到")
        with open(path, "r") as fp:
            for line in fp:
                self.write(line)
        self.finish("")

class LogsApiHandler(WebSocketHandler):

    def allow_draft76(self):
        return True

    def check_origin(self, origin):
        return True

    def on_message(self, message):
        path = os.path.join("datas/logs", message)
        if not os.path.isfile(path):
            return
        if not hasattr(self, "_fp") or self._fp.closed:
            self._fp = open(path, encoding="gbk")
        while 1:
            line = self._fp.readline()
            if not line:
                break
            self.write_message(line)
        IOLoop.instance().add_timeout(time() + 0.5, self.sendLine)

    def on_close(self):
        try:
            self._fp.close()
        except:
            pass

    def sendLine(self):
        if hasattr(self, "_fp") and not self._fp.closed:
            line = self._fp.readline()
            if line:
                self.write_message(line)
            IOLoop.instance().add_timeout(time() + 0.5, self.sendLine)

headers = [
    (r"/api/task/modify", ModifyHandler),  # 任务修改
    (r"/api/task/control", ControlHandler),  # 任务控制
    (r"/api/tasks", IndexHandler),  # 查询任务
    (r"/api/restart", RestartHandler),  # 重启调度器
    (r"/api/logs", LogsApiHandler),  # 日志
    (r"/api/qrcode", QrcodeHandler),  # 二维码
    (r"/api/wechat/send", WeChatHandler),  # 微信消息发送
    (r"/api/email", EmailIdHandler),  # 邮件控制id
    (r"/logs/mms", LogMmsHandler),  # 彩信发送日志
#     (r"/logs/(.*)", LogsHandler),  # 日志
    (r"/logs/(.*)", LogsHandler, {"path":"datas/logs/"}),
    (r"/live", LiveHandler),  # 判断本服务是否启动
    (r"/.*", IndexHandler)  # 首页
]
