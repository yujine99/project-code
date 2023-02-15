
# _*_ coding: utf-8 _*_
# @Time   : 2022/3/15 下午 03:30
# @Author : Carolyn_Yu
# @FileName : 设备异常自动邮件通知.py
# @Software : PyCharm
# Change List:
# 將郵件人員添加到c:\PhoneNumListTemplet.xls；將webdriver改到c:\\temp\\webdriver下；取消打印web错误
# 20取消打印web错误;delete pass and fail log data; add error data to Error_Backup.csv;修改MailAddressListTemplet.xls加副本
# 2022/4/19 float比较大小;
# 2022/4/20 delete empty dir
# 2022/4/21 如果没有mail address栏程式会直接关闭，修改check Mail address
# 2022/4/27 正则表达式check mail address合法性
# 2022/4/2901 删除空文件夹时不显示文件夹内容
# 2022042902 Line栏值如果是test或TEST,不做数据分析直接删除
# 2022050401 subdir_del_list里删除LOG_PATH + r'\line2\123',因为4/28文件删除不掉
# 2022050501 修改mail subject and content
# 2022050502 修改mail subject and content
# 2022050502 修改提示语，请不要锁屏
# 2022050502 修改ftp  \\172.22.248.151\jdm1 tpt04\Laptop C Cover,移除line1/123
# 20220712 屏蔽语音
# 20220914 更换Chrome 浏览器
# 20221207 多mail账户指定发件人
# 2022121606 检查设定的发件人与outlook读出来的是否一致
# 20230116 添加ftp文件夹
# 20230117 取消短信


import sys
import os
import re
from pathlib import Path
import csv
import os.path
import time
from datetime import date
from win32com.client import Dispatch
import win32gui, win32api, win32con
import threading
import pythoncom
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from xlrd import open_workbook

# set var
TEMP_PATH = r'c:\temp_file'
READ_FILE = TEMP_PATH + "\\" + 'read_file.txt'
ERROR_VALUE = TEMP_PATH + "\\" + 'error_value.txt'
ERROR_SOUND = TEMP_PATH + "\\" + 'error_sound.txt'
ERROR_MSG = TEMP_PATH + "\\" + 'error_msg.txt'
ERROR_TEMP_BACKUP = TEMP_PATH + "\\" + 'error_temp_back.csv'
LOG_PATH = r'\\172.22.248.151\jdm1 tpt04\Laptop C Cover'  # 設備數據文件路徑
PHONE_PATH = r'c:\temp\PhoneNumListTemplet.xls'  # 電話號碼文件路徑
MAIL_PATH = r'c:\temp\MailAddressListTemplet.xls'  # 邮件地址文件路徑
ERROR_BACKUP = TEMP_PATH + "\\" + 'Error_Backup.csv'  # 异常数据备份文件路徑
PROGRAM_LOG = TEMP_PATH + "\\" + 'program.log'
EX_MAILTO = ''  # 收件者
EX_MAILCC = ''  # 副本
# subdir_del_list = [LOG_PATH + r'\line1\123', LOG_PATH + r'\line1\320', LOG_PATH + r'\line1\321',
#                    LOG_PATH + r'\line1\322',
#                    LOG_PATH + r'\line2\123', LOG_PATH + r'\line2\326', LOG_PATH + r'\line2\327',
#                    LOG_PATH + r'\line2\328']
subdir_del_list = [LOG_PATH + r'\line1\320', LOG_PATH + r'\line1\321', LOG_PATH + r'\line1\322',
                   LOG_PATH + r'\line2\326', LOG_PATH + r'\line2\327',
                   LOG_PATH + r'\line2\328', LOG_PATH + r'\line2\331', ]


def del_emp_dir(path):
    today = date.today()
    today_str = today.strftime('%Y_%m_%d')
    for (root, dirs, files) in os.walk(path):
        for item in dirs:
            dirname = os.path.join(root, item)
            if (not os.listdir(dirname)) and os.path.basename(dirname) != today_str:  # 非今日空文件夹
                # print(os.path.basename(dirname))
                # print('Today is :' + today_str)
                # print('空文件夹：' + dirname)
                os.rmdir(dirname)
                print('移除非今日空目录: ' + dirname)
            else:
                pass
                # print(dirname, os.listdir(dirname))   # 显示文件夹内容


def write_row_data_csv(file, header, row):
    if not os.path.exists(file):
        print('There have not %s and will create it' % file)
        with open(file, 'a', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(header)
            writer.writerow(row)
            f.close()
    else:
        print('There have %s and will write data' % file)
        with open(file, 'a', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(row)
            f.close()


def read_mail_addr(mail_adr, mail_cc):
    while os.path.exists(MAIL_PATH):
        break
    else:
        print("No Mail address file, please insert 'MailAddressListTemplet.xls' to " + MAIL_PATH)
        sleep(1)
        read_mail_addr(mail_adr, mail_cc)
    wb = open_workbook(MAIL_PATH)
    ws = wb.sheet_by_name('Sheet1')
    if ws.ncols == 0 or ws.nrows == 1:
        print('There have not mail address data, please insert!')
        sleep(1)
        read_mail_addr(mail_adr, mail_cc)
    else:
        regex = re.compile(r'([A-Za-z0-9]+[-._])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+')
        for i in range(1, ws.nrows):  # 打印一列表格的内容
            if ws.ncols == 1:
                if re.fullmatch(regex, ws.cell(i, 0).value):
                    mail_adr = mail_adr + ws.cell(i, 0).value + ';'
                    mail_cc = ''
                else:
                    print('无效的email地址' + ws.cell(i, 0).value)
                    sleep(1)
                    read_mail_addr(mail_adr, mail_cc)
            else:
                if re.fullmatch(regex, ws.cell(i, 0).value):
                    mail_adr = mail_adr + ws.cell(i, 0).value + ';'
                else:
                    print('无效的email地址' + ws.cell(i, 0).value)
                    sleep(1)
                    read_mail_addr(mail_adr, mail_cc)
                if ws.cell(i, 1).value != '':
                    if re.fullmatch(regex, ws.cell(i, 1).value):
                        mail_cc = mail_cc + ws.cell(i, 1).value + ';'
                    else:
                        print('无效的email地址' + ws.cell(i, 1).value)
                        sleep(1)
                        read_mail_addr(mail_adr, mail_cc)
    return mail_adr, mail_cc


def check_webdriver():
    while os.path.exists(r'C:\TEMP\msedgedriver.exe'):
        print('瀏覽器driver文件路徑：' + r'C:\TEMP\msedgedriver.exe' + ', msedgedriver file is OK!')
        break
    else:
        print("No msedgedriver file, please insert 'msedgedriver.exe' to " + TEMP_PATH)
        sleep(1)
        check_webdriver()


def send_mail(sub, content, attach1, attach2):
    pythoncom.CoInitialize()
    ol_format_html = 2
    # OLFormatPlain = 1
    # olFormatRichText = 3
    # olFormatUnspecified = 0
    ol_mail_item = 0x0
    obj = Dispatch("Outlook.Application")

    send_account = None
    # 选择要使用的邮箱账户
    set_send_account = 'Carolyn_Yu@pegatroncorp.com'
    print('-------------------------------------')
    print('设定的邮件发件人与outlook读出来的必须一致！')
    print('添加outlook账号时，请将账号首字母大写，如 Carolyn_Yu@pegatroncorp.com！')
    print('-------------------------------------')
    print('程式设定的邮件发件人是: %s' % set_send_account)
    print('Outlook读出来的账号如下:')
    for account in obj.Session.Accounts:
        print(account)
        if account.DisplayName == set_send_account:  # 需与读出来的一致，否则报错
            send_account = account
            break

    new_mail = obj.CreateItem(ol_mail_item)
    new_mail._oleobj_.Invoke(*(64209, 0, 8, 0), send_account)  # 指定发件人
    new_mail.Subject = sub
    new_mail.BodyFormat = ol_format_html
    # new_mail.HTMLBody = "<h1>I am a title</h1><p>I am a paragraph</p>"
    new_mail.HTMLBody = content
    new_mail.To = EX_MAILTO

    # carbon copies and attachments (optional)
    new_mail.CC = EX_MAILCC
    # new_mail.BCC = "Hong-ze_Wang@pegatroncorp.com"
    new_mail.Attachments.Add(attach1)
    new_mail.Attachments.Add(attach2)

    # open up in a new window and allow review before send
    # new_mail.display()

    # or just use this instead of .display() if you want to send immediately
    new_mail.Send()
    pythoncom.CoUninitialize()  # 多线程需要pythoncom初始及关闭，否则报错
    print('Mail sending complete!')
    #  如何禁用系统弹窗: OutLook选项——信任中心——信任中心设置——编程访问——从不向我发出可疑活动警告


def text_to_sound(voice):
    pythoncom.CoInitialize()
    with open(voice, 'r', encoding='utf-8') as fvs:
        reader = fvs.read()
        # print(reader)
        speaker = Dispatch("SAPI.SpVoice")
        speaker.Speak(reader)
        print('The sound playback complete!')
        pythoncom.CoUninitialize()  # 多线程需要pythoncom初始及关闭，否则报错


def send_msg(msg, phone_path):
    url = 'http://epsz5.sz.pegatroncorp.com/Masts/Modules/MessageSend/MessageSendByExcel.aspx'
    # edge浏览器
    # svs = Service(r'C:\TEMP\msedgedriver.exe')
    # driver = webdriver.Edge(service=svs)  # 打包換電腦后將msedgedriver.exe複製到相應路徑能正常測試
    # # driver = webdriver.Edge()    # 打包換電腦會顯示webdirver PATH錯誤
    # Chrome浏览器
    svs = Service(r'C:\TEMP\chromedriver.exe')
    driver = webdriver.Chrome(service=svs)
    driver.maximize_window()  # 浏览器最大化
    driver.implicitly_wait(10)  # 隐性等待10s
    driver.get(url)
    phone_xlsx = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_fuPhoneNum')
    phone_xlsx.send_keys(phone_path)
    insert_btn = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_btnExport')
    insert_btn.click()
    locator = (By.ID, 'ctl00_ContentPlaceHolder1_txtMessage')
    # noinspection PyBroadException 关闭捕获的异常过于宽泛报警
    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located(locator))  # 显性等待30s
        msg_text = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_txtMessage')
        msg_text.send_keys(msg)
    except Exception as e:
        print(e)
    sleep(5)
    msg_result = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_txtMessageResult')
    msg_result.click()
    send_btn = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_btnSend')
    send_btn.click()
    sleep(5)
    print('Message send OK!')
    driver.close()


def check_phone(phone_path):
    while os.path.exists(phone_path):
        print('電話號碼文件路徑：' + phone_path + ', PhoneNumber file is OK!')
        break
    else:
        print("No PhoneNumber file, please insert 'PhoneNumListTemplet.xls' to " + phone_path)
        sleep(1)
        check_phone(phone_path)


def if_exist_dir_f(file_path, file_name, first_str):  # 判断是否存在文件目录及文件
    if os.path.exists(file_path):
        # 该文件路径存在"
        if not os.path.exists(file_name):
            message = 'Sorry, I cannot find the "%s" path and I create it.' % file_name
            print(message)
            with open(file_name, 'w', encoding='utf-8') as f0:
                # f0.write('The all compared files. \n')  # 创建第一个文件
                f0.write(first_str + '\n')  # 创建第一个文件
    else:
        os.mkdir(file_path)  # 路径不存在，创建文件目录
        if_exist_dir_f(file_path, file_name, first_str)  # 递归


def create_ok_file(file_name, file_content):
    basename = Path(file_name).stem
    ok_name = TEMP_PATH + '\\' + basename + '_' + file_content + '.ok'
    with open(ok_name, 'a', encoding='utf-8') as fok:
        fok.write(file_name + ' ' + file_content + ' send complete!\n')
        print(ok_name + ' create complete!')


def chk_ok_file(file_name, file_content):
    basename = Path(file_name).stem
    ok_name = TEMP_PATH + '\\' + basename + '_' + file_content + '.ok'
    if os.path.exists(ok_name):
        print('There have ' + ok_name + ' file.')
        return True
    else:
        print('There have not ' + ok_name + ' file. Will send ' + file_content + '!')
        return False


def data_comp(filename):
    with open(filename, 'r') as f:
        reader = csv.reader(f)
        header = next(reader)
        for row in reader:
            # print(row[2])
            if row[8] == 'test' or row[8] == 'TEST':
                print('Line column value is %s,it is a debug file, will delete!' % row[8])
                sleep(2)
                break
            if float(row[4]) <= float(row[2]) <= float(row[3]):
                print('%s(%s, %s) Value is OK!' % (row[2], row[3], row[4]))
            else:
                write_row_data_csv(ERROR_TEMP_BACKUP, header, row)
                # mail 信息
                error_info = 'Line-' + row[8] + ' Station-' + row[6] + ' DeviceID-' + row[5] + ' ' + \
                             row[1] + ' Value is %s(%s, %s), exceeded Limits!' % (row[2], row[4], row[3])
                with open(ERROR_VALUE, mode='a', encoding='utf-8') as fv:
                    fv.write(''.join(['%s\n' % error_info]))
                    fv.close()
                print(error_info)
                # 短信信息
                error_msg = '設備異常：' + row[8] + '綫 ' + row[6] + '站' + ' ' + row[1] + '值%s(%s, %s)超標!' % (
                    row[2], row[4], row[3])
                with open(ERROR_MSG, mode='a', encoding='utf-8') as fmsg:
                    fmsg.write(''.join(['%s\n' % error_msg]))
                    fmsg.close()
                # print(error_msg)
                # 语音信息
                error_sound = '設備異常：' + row[8] + '綫 ' + row[6] + ' Station ' + 'DeviceID ' + row[5] + ' ' + row[
                    1] + '值超標!'
                with open(ERROR_SOUND, mode='a', encoding='utf-8') as fsd:
                    fsd.write(''.join(['%s\n' % error_sound]))
                    fsd.close()
                # print(error_sound)

        if os.path.exists(ERROR_TEMP_BACKUP):
            if not chk_ok_file(filename, 'write_backup'):
                with open(ERROR_TEMP_BACKUP, 'r', encoding='utf-8') as fet:
                    reader_e = csv.reader(fet)
                    header_e = next(reader_e)
                    for row in reader_e:
                        # print(row)
                        write_row_data_csv(ERROR_BACKUP, header_e, row)
                    fet.close()
                    create_ok_file(filename, 'write_backup')  # 創建备份错误完成的標誌文件
                    os.remove(ERROR_TEMP_BACKUP)
            else:
                os.remove(ERROR_TEMP_BACKUP)
                print('Have already write error data to backup file and will not write again!')

        if os.path.exists(ERROR_VALUE):  # 判斷錯誤信息文件是否存在
            if not chk_ok_file(filename, 'mail'):  # 判斷是否已經發送過，如果已發送過就不再發送
                with open(ERROR_VALUE, 'r', encoding='utf-8') as f2:
                    content = f2.read()
                    f2.close()
                send_mail('設備超標: ' + content, filename + ' ' + content, filename, ERROR_VALUE)
                create_ok_file(filename, 'mail')  # 創建發送過的標誌文件
                os.remove(ERROR_VALUE)
                # print('error_value.txt have been deleted')
            else:
                os.remove(ERROR_VALUE)
                print('Have already sent mail and will not send again!')


        # 发送短信
        # if os.path.exists(ERROR_MSG):
        #     if not chk_ok_file(filename, 'msg'):
        #         with open(ERROR_MSG, 'r', encoding='utf-8') as fmsg:
        #             msg_content = fmsg.read()
        #             fmsg.close()
        #         send_msg(msg_content, PHONE_PATH)
        #         create_ok_file(filename, 'msg')
        #         os.remove(ERROR_MSG)
        #         # print('error_msg.txt have been deleted')
        #     else:
        #         os.remove(ERROR_MSG)
        #         print('Have already sent message and will not send again!')


        # 语音通知
        # if os.path.exists(ERROR_SOUND):
        #     if not chk_ok_file(filename, 'sound'):
        #         with open(ERROR_SOUND, 'r', encoding='utf-8') as fsd:
        #             sud_content = fsd.read()
        #             fsd.close()
        #             print(sud_content)
        #         text_to_sound(ERROR_SOUND)
        #         text_to_sound(ERROR_SOUND)
        #         text_to_sound(ERROR_SOUND)
        #         create_ok_file(filename, 'sound')
        #         os.remove(ERROR_SOUND)
        #         # print('error_sound.txt have been deleted')
        #     else:
        #         os.remove(ERROR_SOUND)
        #         print('Have already sent sound and will not send again!')

        if os.path.exists(TEMP_PATH + '\\' + Path(filename).stem + '_' + 'mail' + '.ok'):
            os.remove(TEMP_PATH + '\\' + Path(filename).stem + '_' + 'mail' + '.ok')
        # if os.path.exists(TEMP_PATH + '\\' + Path(filename).stem + '_' + 'msg' + '.ok'):
        #     os.remove(TEMP_PATH + '\\' + Path(filename).stem + '_' + 'msg' + '.ok')
        # if os.path.exists(TEMP_PATH + '\\' + Path(filename).stem + '_' + 'sound' + '.ok'):
        #     os.remove(TEMP_PATH + '\\' + Path(filename).stem + '_' + 'sound' + '.ok')
        if os.path.exists(TEMP_PATH + '\\' + Path(filename).stem + '_' + 'write_backup' + '.ok'):
            os.remove(TEMP_PATH + '\\' + Path(filename).stem + '_' + 'write_backup' + '.ok')


# data_comp(r'G:\PycharmProjects\pythonProject\设备异常自动邮件通知\20220323\221431970000005_2022_03_17_12_51_39.csv')

# Thread1
def check_log(path):
    while True:
        try:
            for i in subdir_del_list:
                del_emp_dir(i)
            if os.path.exists(ERROR_TEMP_BACKUP):
                os.remove(ERROR_TEMP_BACKUP)
            else:
                pass
            if os.path.exists(ERROR_VALUE):
                os.remove(ERROR_VALUE)
            else:
                pass
            if os.path.exists(ERROR_SOUND):
                os.remove(ERROR_SOUND)
            else:
                pass
            if os.path.exists(ERROR_MSG):
                os.remove(ERROR_MSG)
            else:
                pass
            for root, dirs, files in os.walk(path):
                for name in files:
                    file_rel_path = os.path.join(root, name)
                    # file_mtime = os.stat(file_rel_path).st_mtime
                    # a = time.time() - file_mtime
                    # if time.time() - file_mtime < 600:  # 10分钟
                    print(file_rel_path)
                    # print('Duration from file creation to test is ' + str(a) + 's.')
                    if_exist_dir_f(TEMP_PATH, READ_FILE, 'The all compared files.')
                    with open(READ_FILE, encoding='utf-8') as f1:
                        for line in f1:
                            if file_rel_path.strip() in line.strip():
                                # 按行读取read_file.tx，如果该行包含file_rel_path，则表明已经比较过，写have_file.txt
                                with open(r'C:\temp_file\have_file.txt', 'w', encoding='utf-8') as fh:
                                    fh.write('The file %s is in read_file.txt' % file_rel_path)
                            else:
                                pass
                        if os.path.exists(r'C:\temp_file\have_file.txt'):  # 有have_file.txt表明已经compare，退出
                            print('The file %s has already been compared' % file_rel_path)
                            os.remove(r'C:\temp_file\have_file.txt')
                            os.remove(file_rel_path)
                            print('The file %s has been deleted' % file_rel_path)
                        else:
                            data_comp(file_rel_path)  # 没有have_file.txt表明未曾比较过，compare并将文件名写入read_file.txt
                            print('Will write compared file %s to %s' % (file_rel_path, READ_FILE))
                            with open(READ_FILE, 'a', encoding='utf-8') as fr:
                                fr.write(''.join(['%s\n' % file_rel_path]))
                            os.remove(file_rel_path)
                            print('The file %s has been deleted' % file_rel_path)

                    # else:
                    #     pass
                    #     # print('No new file!')
        except Exception as e:
            print(e)
            print('执行接下去的代码')


# check_log(log_path)

# Thread3
def mouse_click(find_window, mark1):
    while True:
        try:
            time.sleep(10)
            if os.path.exists(mark1):
                # # 获取所有窗口句柄
                # hwnd_title = {}
                #
                # def get_all_hwnd(hwnd, mouse):
                #     if (win32gui.IsWindow(hwnd)
                #             and win32gui.IsWindowEnabled(hwnd)
                #             and win32gui.IsWindowVisible(hwnd)):
                #         hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})
                #
                # win32gui.EnumWindows(get_all_hwnd, 0)
                # for h, t in hwnd_title.items():
                #     if t:
                #         print(h, t)

                # 置顶窗口
                print("置顶窗口")
                # 窗口需要正常大小且在后台，不能最小化
                hwnd = win32gui.FindWindow(None, find_window)
                if (win32gui.IsWindow(hwnd)
                        and win32gui.IsWindowEnabled(hwnd)
                        and win32gui.IsWindowVisible(hwnd)):

                    # 激活显示窗口，使其成为置顶活动窗口
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
                    # win32gui.SetForegroundWindow(hwnd)
                    # 置顶
                    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                          win32con.SWP_NOMOVE | win32con.SWP_NOACTIVATE | win32con.SWP_NOOWNERZORDER | win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE)
                    # 取消置顶
                    # win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,win32con.SWP_SHOWWINDOW|win32con.SWP_NOSIZE|win32con.SWP_NOMOVE)

                    # 获取窗口的位置信息
                    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                    x = int(left + (right - left) // 3.7)
                    y = int(top + (bottom - top) // 6 * 5.2)
                    print(x, y)
                    time.sleep(1)
                    win32api.SetCursorPos([x, y])
                    time.sleep(0.1)
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                    time.sleep(0.1)
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                    time.sleep(0.3)
                    print('Finded Outlook msg window and clicked!')
                else:
                    print('No Outlook msg window!')
            else:
                pass
                # print('No new mail send!')
        except Exception as e:
            print(e)


# mouse_click('Microsoft Outlook')

# # Thread2
# import pyautogui
# def capture_mail_window(mark1):
#     while True:
#         try:
#             if os.path.exists(mark1):
#                 time.sleep(10)
#                 pyautogui.click(698, 501, 2, 0.25, button='left')
#                 pyautogui.FAILSAFE = True
#             else:
#                 pass
#                 # print('No new mail send!')
#         except Exception as e:
#             print(e)
#
#
# # capture_mail_window(r'C:\temp_file\error_value.txt')


# main process part:
if_exist_dir_f(TEMP_PATH, PROGRAM_LOG, 'The log file.')  # 判断log file文件及目录是否存，不存在就创建


# 定义print同时输出屏幕及文件log
class Logger(object):
    def __init__(self, filename="log.txt"):
        self.terminal = sys.stdout
        self.log = open(filename, "a", encoding='utf-8')

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
        self.log.flush()  # 缓冲区的内容及时更新到log文件中

    def flush(self):
        pass


# path = os.path.abspath(os.path.dirname(__file__))
# typea = sys.getfilesystemencoding()
sys.stdout = Logger(PROGRAM_LOG)  # 定义print同时输出屏幕及program.log文件, 之后用print输出的就既在屏幕上又在log文件里

# 提示信息
print('*' * 83 + '\n' + '*' * 6 + ' 注意事项：请打开Outlook并将窗口最小化! 请不要锁屏（可以关闭屏幕）!              ' + '*' * 6)
print('*' * 6 + " 請將'msedgedriver.exe'文件放到'C:\\TEMP'文件夹下!                         " + '*' * 6)
print('*' * 6 + " 請將郵件地址文件'MailAddressListTemplet.xls'放到'C:\\TEMP'文件夹下!         " + '*' * 6)
print('*' * 6 + " 請將電話號碼'PhoneNumListTemplet.xls'并放到'C:\\TEMP'文件夹下!              " + '*' * 6)
print('*' * 6 + " 異常數據備份文件：'C:\\temp_file\\Error_Backup.csv'                        " + '*' * 6)
print('*' * 6 + " 添加outlook账号时，请将账号名首字母大写，如 Carolyn_Yu@pegatroncorp.com！    " + '*' * 6 + '\n' + '*' * 83)

print(time.asctime())  # 当前日期时间
print('設備數據文件路徑：' + LOG_PATH)
check_webdriver()  # 瀏覽器driver文件是否存在
check_phone(PHONE_PATH)  # 检查电话文件是否存在
print('郵件地址文件路徑：' + MAIL_PATH + ', Mail address file is OK!')
mail_list = read_mail_addr(EX_MAILTO, EX_MAILCC)
EX_MAILTO = mail_list[0]  # 读取收件者邮件地址
EX_MAILCC = mail_list[1]  # 读取副本邮件地址
print('收件者: ' + EX_MAILTO)  # 收件者
print('副 本: ' + EX_MAILCC)  # 副本
time.sleep(10)

if __name__ == '__main__':
    p1 = threading.Thread(target=check_log, args=(LOG_PATH,))
    # p2 = threading.Thread(target=capture_mail_window, args=(r'C:\temp_file\error_value.txt',))
    p3 = threading.Thread(target=mouse_click, args=('Microsoft Outlook', ERROR_VALUE,))
    p1.start()
    # p2.start()
    p3.start()
