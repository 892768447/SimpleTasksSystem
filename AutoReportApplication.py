#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月1日
@author: Irony."[讽刺]
@site: http://alyl.vip, http://orzorz.vip, https://coding.net/u/892768447, https://github.com/892768447
@email: 892768447@qq.com
@file: AutoReportApplication
@description: 
'''
import base64
import uuid

from tornado.web import Application

from AutoReportHandlers import headers  # @UnresolvedImport


__Author__ = "By: Irony.\"[讽刺]\nQQ: 892768447\nEmail: 892768447@qq.com"
__Copyright__ = "Copyright (c) 2017 Irony.\"[讽刺]"
__Version__ = "Version 1.0"


class AutoReportApplication(Application):

    @property
    def cookie_secret(self):
        return base64.b64encode(uuid.uuid4().bytes + uuid.uuid4().bytes)

    @property
    def scheduler(self):
        return self._scheduler
    
    @property
    def logger(self):
        return self._logger

    def __init__(self, logger, scheduler):
        self._logger = logger
        self._scheduler = scheduler
        settings = {
            "cookie_secret": self.cookie_secret,
            "xsrf_cookies": False,
            "gzip": True,
            "template_path": "datas/web/template",
            "static_path": "datas/web/static",
            "debug": False,
        }
        super(AutoReportApplication, self).__init__(headers, **settings)
        # 关闭模版缓存
        self.settings.setdefault("compiled_template_cache", False)
