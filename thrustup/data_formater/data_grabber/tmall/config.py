#!/usr/bin/env python
# coding: utf-8


_DBUSER = "root" # 数据库用户名
_DBPASS = "root" # 办公数据库密码
# _DBPASS = "mysql" # 个人PC数据库密码
_DBHOST = "127.0.0.1" # 数据库地址
_DBNAME = "tmall" # 数据库名称



class rec: pass

rec.database = 'mysql://%s:%s@%s/%s' % (_DBUSER, _DBPASS, _DBHOST,_DBNAME)
rec.description = u"my blog"
rec.url = 'http://www.demo.com'
rec.paged = 8
rec.archive_paged = 20
rec.admin_username = 'binosun'
rec.admin_email = ''
rec.admin_password = ''
rec.default_timezone = "Asia/Shanghai"
