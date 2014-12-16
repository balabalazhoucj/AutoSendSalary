﻿#-*- coding: utf-8 -*-
#!/usr/bin/python
#filename : rt_compsemail.py
#author   : doCooler

import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import COMMASPACE, formatdate
import mimetypes
import os

class Ak_mailGen:
    def __init__(self):
        self.mailContent = MIMEMultipart()
        self.attatchment = ''
        self.bodyFile    = ''
    
    def set_mail_attach(self, filename=[]):
        self.attachment = filename

    def set_mail_head(self, to_addr,from_addr,subject):
        self.mailContent['To']      = to_addr
        self.mailContent['From']    = from_addr
        self.mailContent['Subject'] = subject
        self.mailContent['Date']    = formatdate(localtime=True)


    def set_body_file(self, filename):
        self.bodyFile = filename

    def gen_mail(self):
        #coding = gbk
        import codecs
        file_content = open(self.bodyFile, 'rb').read()
        file_content = file_content.decode("gbk")
        #print (file_content)
        body = MIMEText(file_content, 'plain' , 'UTF-8')
        #body["Content-Type"] = 'plain/text'
        self.mailContent.attach(body)

        for file in self.attachment:
            part = MIMEText(open(file, 'rb').read(), 'base64', "UTF-8")
            part["Content-Type"] = 'application/octet-stream'
            part["Content-Disposition"] = 'attachemnt; filename =  ' + "工资条.xls"
            self.mailContent.attach(part)


    def get_mail_content(self):
        return self.mailContent.as_string()



