#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年7月2日
@author: Irony."[讽刺]
@site: http://alyl.vip, http://orzorz.vip, https://coding.net/u/892768447, https://github.com/892768447
@email: 892768447@qq.com
@file: AutoReportScheduler
@description: 
'''

__Author__ = "By: Irony.\"[讽刺]\nQQ: 892768447\nEmail: 892768447@qq.com"
__Copyright__ = "Copyright (c) 2017 Irony.\"[讽刺]"
__Version__ = "Version 1.0"

# 调度器
Scheduler = None

# 邮件ID
EmailIds = None

# jvm是否启动
JVMStarted = False

# 绿网Session
GreenCookie = None

# 微信机器人实例
WeixinBot = None

# 彩信日志路径
PathNowSMS = "C:/Program Files/NowSMS"
# 彩信日志文件名格式化
NameNowSMS = "MMSC-%Y%m%d.LOG"
