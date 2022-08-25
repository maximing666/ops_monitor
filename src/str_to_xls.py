import re
import xlrd    # 读取excel
import xlwt    # 写excel
from xlutils.copy import copy
import os
import getmail

def get_mail():
    my_mail = getmail.main()
    # mailheader = my_mail[0] # 获取邮件头信息
    # print(mailheader)
    mailtext = my_mail[1]   # 获取邮件内容
    return mailtext

def main():
    my_mailtext = get_mail()
    my_mailtext_l = my_mailtext.split()


    # 修改excle:
    if os.path.exists('opscheck_new.xls'):  # 新文件
        os.remove('opscheck_new.xls')
    excelfile = 'opscheck.xls'  # 原文件

    # 样式一
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    borders.bottom_colour = 0x3A       # 深绿色
    style = xlwt.XFStyle()
    style.borders = borders

    # 样式二
    borders2 = xlwt.Borders()
    borders2.left = 1
    borders2.right = 1
    borders2.top = 1
    borders2.bottom = 1
    borders2.bottom_colour = 0x3A         # 深绿色
    style2 = xlwt.XFStyle()
    style2.borders = borders
    style2.font.colour_index = 0x0A                    # 红色

    # 上下左右居中
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style.alignment = alignment
    style2.alignment = alignment

    work = xlrd.open_workbook(excelfile, formatting_info=True)
    old_content = copy(work)
    ws = old_content.get_sheet(0)
    l_mail = [my_mailtext_l[3]]       # 获取邮件中有用数据，将此数据填入新xls空白处
    for i in my_mailtext_l[5::3]:
        l_mail.append(i)
    for i in l_mail:                      # 精简内容
        if i.find('正常') >= 0:
            l_mail[l_mail.index(i)] = '正常'
        elif i.find('异常') >= 0:
            l_mail[l_mail.index(i)] = '异常'
    sheet = work.sheet_by_name('巡检日志')  # 获取该sheet页
    ret = sheet.row(2)  # 获取第3行数据
    l_blank = []        # 获取原xls文件中第2行中空白内容的列号
    cls_r = 0
    for i in ret:
        if i.value == '':
            l_blank.append(cls_r)
        cls_r += 1
    for i in l_blank:
        if l_mail[i] == '异常':    # 如果异常，则字体颜色为红色
            ws.write(2, i, l_mail[i], style2)
        else:
            ws.write(2, i, l_mail[i], style)

    old_content.save('opscheck_new.xls')
