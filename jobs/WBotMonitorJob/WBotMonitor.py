#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月19日
@author: Irony."[讽刺]
@site: http://alyl.vip, http://orzorz.vip, https://coding.net/u/892768447, https://github.com/892768447
@email: 892768447@qq.com
@file: jobs.WBotMonitorJob.WBotMonitor
@description: 
'''
import os
import re
from threading import Thread

from wxpy.api.bot import Bot
from wxpy.ext.tuling import Tuling

import AutoReportGlobals  # @UnusedImport @UnresolvedImport
from AutoReportUtil import sendMms  # @UnresolvedImport


__Author__ = "By: Irony.\"[讽刺]\nQQ: 892768447\nEmail: 892768447@qq.com"
__Copyright__ = "Copyright (c) 2017 Irony.\"[讽刺]"
__Version__ = "Version 1.0"

MMSPATTERN = re.compile("@.*\s(日报发送),([0-9a-zA-Z]{6})($|[\n|\r|\-*|<]|,([0-9,]+))")

class WeixinBot(Thread):

    def __init__(self, logger, groups):
        super(WeixinBot, self).__init__()
        AutoReportGlobals.WeixinBot = self
        self.logger = logger
        self.groups = groups
        self._groups = []
        self.path = "datas/tmps/weixin"
        self.cachePath = self.path + "/wx.pkl"
        self.qrcodePath = "datas/web/static/qr.png"
        self.puidPath = self.path + "/wx_puid.pkl"
        os.makedirs(self.path, exist_ok=True)

    def onQrCall(self, uuid, status, qrcode):
        self.logger.debug("uuid: %s, status: %s, qrcode: %s" % 
                          (uuid, status, len(qrcode)))
        with open(self.qrcodePath, "wb") as fp:
            fp.write(qrcode)
        self.logger.debug("save qrcode ok")

    def onLogin(self, *args, **kwargs):
        try:
            os.unlink(self.qrcodePath)
        except Exception as e:
            self.logger.warn(str(e))

    def onLogout(self, *args, **kwargs):
        self._groups.clear()
        AutoReportGlobals.WeixinBot = None
        self.logger.debug("the weixin bot is logout")
    
    def sendText(self, text):
        try:
            for group in self._groups:
                group.send(text)
        except Exception as e:
            self.logger.warn(str(e))
            return -1
        return 1
    
    def sendImage(self, image):
        try:
            if not os.path.isfile(image):
                self.logger.warn("%s 不是文件或不存在" % image)
                return -1
            for group in self._groups:
                group.send_image(image)
        except Exception as e:
            self.logger.warn(str(e))
            return -1
        return 1
    
    def sendImages(self, images):
        try:
            for image in images:
                if not os.path.isfile(image):
                    self.logger.warn("%s 不是文件或不存在" % image)
                    return -1
                self.sendImage(image)
        except Exception as e:
            self.logger.warn(str(e))
            return -1
        return 1
    
    def dealWithMMS(self, group, results):
        ver, eid, _, whichs = results[0]
        whichs = [w for w in whichs.split(",") if w != ""]
        self.logger.debug("ver:{0}, eid:{1},whichs:{2}".format(ver, eid, whichs))
#         print(AutoReportGlobals.EmailIds)
        emailInfo = AutoReportGlobals.EmailIds.get(eid)
        if not emailInfo:
            return group.send("未找到该缓存 id: {0}".format(eid))
        subject = emailInfo.get("subject")  # 获取主题
        dataDir = emailInfo.get("dataDir")  # 获取数据目录
        peoples = emailInfo.get("formalSend")  # 获取正式发送人
        if ver == "日报发送":
            self.logger.debug("eid:{0}, subject:{1}, dataDir:{2}, peoples:{3}, whichs:{4}".format(eid, subject, dataDir, peoples, whichs))
            result = sendMms(subject, dataDir, peoples, whichs, "")
            group.send("日报控制-" + ("成功" if result == 1 else "失败"))
            del AutoReportGlobals.EmailIds[eid]  # 删除

    def run(self):
        self.bot = Bot(
            cache_path=self.cachePath,  # 会话缓存路径
            # console_qr=True,  # 终端中显示二维码需要pillow
            qr_path=self.qrcodePath,  # 二维码保存路径
            qr_callback=self.onQrCall,  # 二维码回调函数
            login_callback=self.onLogin,  # 登录成功回调函数
            logout_callback=self.onLogout  # 退出时回调函数
        )
        # 图灵机器人
        tuling = Tuling(api_key="11212cbfcbdd532351ebd13118e768e1")

        # puid 是 wxpy 特有的聊天对象/用户ID
        # 居右唯一性
        self.bot.enable_puid(self.puidPath)  # 启用聊天对象的puid属性

        groups = self.bot.groups(update=True)
        self.logger.debug("groups len: %s" % (len(groups)))
        count = 0
        for g in self.groups:
            group = groups.search(g)
            if not group:
                count += 1
                self.logger.warn("无法找到所监控的群")
                continue
            group = group[0]
            self._groups.append(group)
            self.logger.debug("find monitor group")

            # 机器人回复
            @self.bot.register(group)
            def replay(msg):
                if not msg.is_at:
                    return  # 如果没有被at则忽略
                # 彩信日报命令解析
#                 print(msg.text)
                results = MMSPATTERN.findall(msg.text)
#                 print(results)
                if results and len(results) >= 1:  # 发现控制命令
                    self.dealWithMMS(group, results)
                    return
                # 机器人业务
                tuling.do_reply(msg)  # 使用图灵机器人回复
        if count == len(self.groups):  # 如果计数器和所监控的群个数相同
            return self.bot.logout()
        # 自己的群
        try:
            mygroup = groups.search("开发交流")[0]
            @self.bot.register(mygroup)
            def rMygroup(msg):
                tuling.do_reply(msg)  # 使用图灵机器人回复
        except:pass
        # 阻塞当前线程
        self.bot.join()


class WBotMonitor:

    def __init__(self, logger, config, Id):
        self.logger = logger
        self.groups = config.get("群", [])
        self.Id = Id
        self.wbot = None

    def start(self):
        if not self.groups:
            return self.logger.warn("监控的群对象为空")
        if AutoReportGlobals.WeixinBot:
            self.logger.warn("机器人已经实例化")
            job = AutoReportGlobals.Scheduler.get_job(self.Id)
            if job:
                job.pause()  # 暂停该任务
            return
        self.wbot = WeixinBot(self.logger, self.groups)
        self.wbot.start()
