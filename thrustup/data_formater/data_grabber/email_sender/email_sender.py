# -*- coding: utf-8 -*-
import smtplib

import os.path
from email.mime.multipart import MIMEMultipart
from email.mime.multipart import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import email
import json
import base64

class EmailEngine(object):

    def __init__(self,filename=None, config=None):

        self.__load_config(config)
        self.__build_main_msg()
        self.__build_text_msg()
        self.__build_file_msg(filename)
        self.__build_image_msg()

    def __load_config(self,conf):
        if not conf:
            with open("../config/email_config.json","r") as f:
                self.config = json.load(f)
        else:
            self.config = conf
        self.smtp = self.config["smtp"]
        self.account = self.config["account"]
        self.password = base64.decodestring(self.config["password"])
        self.from_email = self.config["from_email"]
        self.to_emails = self.config["to_emails"]
        self.subject = self.config["subject"]
        # self.file_name = self.config["file_name"]
        self.body_name = self.config["body_name"]
        self.image_name = self.config["image_name"]



    def __build_main_msg(self):
        # 构造MIMEMultipart对象做为根容器
        self.main_msg = MIMEMultipart()

        # 设置根容器属性
        self.main_msg['From'] = self.from_email
        self.main_msg['To'] = ",".join(self.to_emails)
        self.main_msg['Subject'] = self.subject
        self.main_msg['Date'] = email.Utils.formatdate()


    def __build_text_msg(self):
        with open(self.body_name,"r") as body:
            self.body = body.read()
        # 构造MIMEText对象做为邮件显示内容并附加到根容
        # text_msg = MIMEText(self.body, 'html', 'utf-8')
        text_msg = MIMEText(self.body, 'html')

        self.main_msg.attach(text_msg)

    def __build_file_msg(self, file_list):
        if file_list:
            for file_path in file_list:
                # 构造MIMEBase对象做为文件附件内容并附加到根容器
                contype = 'application/octet-stream'
                maintype, subtype = contype.split('/', 1)

                ## 读入文件内容并格式化
                try:
                    data = open(file_path, 'rb')
                except Exception as e:
                    return
                file_msg = MIMEBase(maintype, subtype)
                file_msg.set_payload(data.read())
                email.Encoders.encode_base64(file_msg)

                ## 设置附件头
                basename = os.path.basename(file_path)
                file_msg.add_header('Content-Disposition',
                                    'attachment', filename=basename)

                self.main_msg.attach(file_msg)
        else:
            pass

    def __build_image_msg(self):
        image = open(self.image_name, 'rb')
        image_msg = MIMEImage(image.read())
        image.close()

        ## 设置附件头
        basename = os.path.basename(self.image_name)
        image_msg.add_header('Content-Disposition',
                            'attachment', filename=basename)
        self.main_msg.attach(image_msg)


    def send_email(self):
        server = smtplib.SMTP(self.smtp)
        server.login(self.account, self.password)  # 仅smtp服务器需要验证时

        # 得到格式化后的完整文本
        fullText = self.main_msg.as_string()
        try:
            server.sendmail(self.from_email, self.to_emails, fullText)
            # server.sendmail(self.from_email,)
        finally:
            server.quit()
