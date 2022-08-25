#!/usr/bin/python
# -*- coding: UTF-8 -*-
# 不带附件
import smtplib
from email.header import Header
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart

my_sender = '61548681@qq.com'  # 发件人邮箱账号
my_pass = 'kpjyuurcixxncaaj'  # 发件人邮箱密码
to_receiver = 'maxm@cecurs.com'  # 收件人邮箱账号，我这边发送给自己


def mail():
    ret = True
    try:
        msg = MIMEMultipart()
        msg['From'] = formataddr(["小马哥", my_sender])  # 括号里的对应发件人邮箱昵称、发件人邮箱账号
        msg['To'] = formataddr(["中电信用", to_receiver])  # 括号里的对应收件人邮箱昵称、收件人邮箱账号
        msg['Subject'] = Header("巡检img", 'utf-8')  # 邮件的主题，也可以说是标题
        msg.attach(MIMEText('Good luck!', 'plain', 'utf-8'))

        # 构造附件（png格式的图片）
        att = MIMEText(open('data.png', 'rb').read(), 'base64', 'utf-8')
        att["Content-Type"] = 'application/octet-stream'
        att["Content-Disposition"] = 'attachment; filename="data.png"'
        msg.attach(att)

        server = smtplib.SMTP_SSL("smtp.qq.com", 465)  # 发件人邮箱中的SMTP服务器，端口是25
        server.login(my_sender, my_pass)  # 括号中对应的是发件人邮箱账号、邮箱密码
        server.sendmail(my_sender, [to_receiver, ], msg.as_string())  # 括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件
        server.quit()  # 关闭连接
    except Exception:  # 如果 try 中的语句没有执行，则会执行下面的 ret=False
        ret = False
    return ret

def main():
    ret1 = mail()
    if ret1:
        print("邮件发送成功")
    else:
        print("邮件发送失败")



'''
构造不同类型的附件：
# 构造附件1（附件为TXT格式的文本）
att1 = MIMEText(open('text1.txt', 'rb').read(), 'base64', 'utf-8')
att1["Content-Type"] = 'application/octet-stream'
att1["Content-Disposition"] = 'attachment; filename="text1.txt"'
message.attach(att1)

# 构造附件2（附件为JPG格式的图片）
att2 = MIMEText(open('123.jpg', 'rb').read(), 'base64', 'utf-8')
att2["Content-Type"] = 'application/octet-stream'
att2["Content-Disposition"] = 'attachment; filename="123.jpg"'
message.attach(att2)

# 构造附件3（附件为HTML格式的网页）
att3 = MIMEText(open('report_test.html', 'rb').read(), 'base64', 'utf-8')
att3["Content-Type"] = 'application/octet-stream'
att3["Content-Disposition"] = 'attachment; filename="report_test.html"'
'''
