# -*- coding: utf-8 -*-
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

sender = 'jimmyyu@fortune-co.com'
smtpserver = '220.181.97.136'
username = 'jimmyyu@fortune-co.com'
password = 'Xiaoyu822'

def sendTextMailTo(p_receiver,p_subject,p_text):
    msg = MIMEText(p_text,'plain','utf-8')
    msg['Subject'] = p_subject
    msg['from'] = 'jimmyyu@fortune-co.com'
    msg['to'] = ','.join(p_receiver)
    smtp = smtplib.SMTP()
    smtp.connect(smtpserver)
    smtp.login(username, password)
    smtp.sendmail(sender, p_receiver, msg.as_string())
    smtp.quit()

def sendHtmlMailTo(p_receiver,p_subject,p_text):
    msg = MIMEText(p_text,'html','utf-8')
    msg['Subject'] = p_subject
    msg['from'] = 'jimmyyu@fortune-co.com'
    msg['to'] = ','.join(p_receiver)
    smtp = smtplib.SMTP()
    smtp.connect(smtpserver)
    smtp.login(username, password)
    smtp.sendmail(sender, p_receiver, msg.as_string())
    smtp.quit()

def sendMultMailTo(p_receiver,p_subject,p_text,p_type,p_attachList):
    msg = MIMEMultipart()
    msg['Subject'] = p_subject
    msg['from'] = 'jimmyyu@fortune-co.com'
    msg['to'] = ','.join(p_receiver)
    msg.attach(MIMEText(p_text, p_type, 'utf-8'))
    for eachAttatch in p_attachList:
        p_url = eachAttatch[0]
        p_filename = eachAttatch[1]
        with open (p_url,'rb') as f:
            part = MIMEApplication(f.read())
            part.add_header('Content-Disposition', 'attachment', filename=p_filename)
            msg.attach(part)

    smtp = smtplib.SMTP()
    smtp.connect(smtpserver)
    smtp.login(username, password)
    smtp.sendmail(sender, p_receiver, msg.as_string())
    smtp.quit()
