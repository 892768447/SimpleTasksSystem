#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
Created on 2017年8月18日
@author: Irony."[讽刺]
@site: alyl.vip, orzorz.vip, irony.coding.me , irony.iask.in , mzone.iask.in
@email: 892768447@qq.com
@file: jobs.ReportCollectionJob.营销报表提取配置.配置读取测试
@description: 
'''

__version__ = "0.0.1"

import yaml

codes = yaml.load(open("codes.txt").read())
print(codes)

config = yaml.load(open("config.txt").read())
print(config)

cookies = yaml.load(open("cookies.txt").read())
print(cookies)

params = yaml.load(open("params.txt").read())
print(params)

referers = yaml.load(open("referers.txt").read())
print(referers)

urls = yaml.load(open("urls.txt").read())
print(urls)