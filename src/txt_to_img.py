from PIL import ImageGrab
import xlwings as xw
import sys
import time
# 需要运行该脚本的服务器上安装有WPS，安装微软的office程序可能会报错。

# 标准信息输出到日志文件
# log_f = open('log.txt', 'w')
# sys.stderr = log_f
# sys.stdout = log_f


# get_screenshot
def excel_catch_screen(shot_excel, shot_sheetname):
    app = xw.App(visible=False, add_book=False)  # 使用xlwings的app启动，参数visible（表示处理过程是否可视，也就是处理Excel的过程会不会显示出来），add_book（是否打开新的Excel程序，也就是是不是打开一个新的excel窗口）
    wb = app.books.open(shot_excel)  # 打开文件
    sheet = wb.sheets(shot_sheetname)  # 选定sheet
    all = sheet.used_range  # 获取有内容的range
    # print(all.value)
    all.api.CopyPicture()  # 复制图片区域
    sheet.api.Paste()  # 粘贴
    img_name = 'data'
    pic = sheet.pictures[0]  # 当前图片
    pic.api.Copy()  # 复制图片
    img = ImageGrab.grabclipboard()  # 获取剪贴板的图片数据
    img.save(img_name + ".png")  # 保存图片
    pic.delete()  # 删除sheet上的图片
    wb.close()  # 不保存，直接关闭
    app.quit()


def main():
    time.sleep(5)
    excel_catch_screen('opscheck_new.xls', '巡检日志')


