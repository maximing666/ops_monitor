# __*__ coding:utf-8 __*__
import poplib
from email.parser import Parser
from email.header import decode_header
# from email.utils import parseaddr
# import chardet
# import quopri
import time

# 输入邮件地址，口令和POP3服务器地址
email = 'maxm@cecurs.com'
password = 'mxm*201804'
pop3_server = 'pop.ym.163.com'

from_my = 'cecyunwei@cecurs.com'
date = time.strftime("%Y-%m=%d+", time.localtime()).replace('-', '年').replace('=', '月').replace('+', '日')
# subject_my1 = date + '__06:10中信信用--北京日常运维巡检'
# subject_my2 = date + '__16:20中信信用--北京日常运维巡检'
subject_my1 = '中电信用--北京日常运维巡检'
# subject_my2 = '日__16:20中信信用--北京日常运维巡检'
# my_header = ''

def conn_mails(email, password, pop3_server):
    # 连接到POP3服务器
    server = poplib.POP3(pop3_server)
    # 调试信息
    # server.set_debuglevel(1)
    # 可选：打印POP3服务器的欢迎文字
    # print(server.getwelcome().decode('utf-8'))

    # 身份认证
    server.user(email)
    server.pass_(password)

    # stat()返回邮件数量和占用空间
    # print('Messages: %s. Size: %s' % server.stat())
    # list()返回所有邮件的编号
    resp, mails, octets = server.list()
    #可以查看返回的列表，类似[b'1 82923',b'2 2184',...]
    # print('mails====:  ', mails)
    index = server.stat()[0]
    return server, index


def getmsg(mailserver, index):
    # 获取最新一封邮件，index为邮件编号
    resp, lines, octets = mailserver.retr(index)
    # lines存储了邮件的原始文本的每一行
    # 可以获得整个邮件的原始文本
    # print('lines:  ', lines)
    msg_context = b'\r\n'.join(lines).decode('utf-8')
    # 解析邮件内容
    msg = Parser().parsestr(msg_context)
    return msg


def header_tmp(msg, indent=0):
    my_header = ''
    if indent == 0:
        for header in ['From', 'To', 'Date', 'Subject']:
            value = msg.get(header, '')
            if value:
                if header == 'Subject':
                    value = decode_str(value)
                    global subject_tmp
                    subject_tmp = value
            if header == 'From':
                global from_tmp
                from_tmp = value
            my_header = my_header + header +':' + value + '\n'
    return from_tmp, subject_tmp, my_header


# indent用于缩进显示:
def print_info(msg, indent=0):
    if (msg.is_multipart()):
        parts = msg.get_payload()
        for n, part in enumerate(parts):
            print('%spart %s' % ('  ' * indent, n))
            print('%s--------------------' % ('  ' * indent))
            print_info(part, indent + 1)
    else:
        content_type = msg.get_content_type()    # 邮件内容的文件类型
        if content_type == 'text/plain' or content_type == 'text/html':  # text/plain是文本，text/html是网页
            content = msg.get_payload(decode=True)  # 返回邮件内容
            charset = guess_charset(msg)            # 返回邮件内容的编码类型，如utf-8。此处有问题？
            if charset:
                content = content.decode('unicode_escape')    #  对邮件内容解密.因为content是unicode编码，所以用unicode_escape解码。
            # print('%sText: %s' % ('  ' * indent, content + '...'))   # 打印邮件内容
        else:
            print('%sAttachment: %s' % ('  ' * indent, content_type))
    return content

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value


def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset


def main():
    my_header = ''
    my_content = ''
    conn = conn_mails(email, password, pop3_server)
    mailserver = conn[0]
    mailnum = conn[1]
    for i in range(mailnum, 0, -1):
        msg_tmp = getmsg(mailserver, i)
        header = header_tmp(msg_tmp, indent=0)
        from_tmp = header[0]
        subject_tmp = header[1]
        if from_tmp == from_my and subject_tmp.find(subject_my1) >= 0:
            my_header = header[2]
            my_content = print_info(msg_tmp, indent=0).strip()
            break
    mailserver.quit()
    # print(my_header)    # 输出邮件首部
    # print(my_content)   # 输出邮件内容
    return my_header, my_content



