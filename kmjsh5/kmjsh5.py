# -*- coding=utf-8 -*-
# @Time     :2019/9/29 10:15
# @Author   :ZhouChuqi
import xlrd, xlwt
import time
from selenium import webdriver
from xlutils.copy import copy
from datetime import datetime
import sys

reload(sys)
sys.setdefaultencoding('utf-8')
import traceback
from selenium.webdriver.common.keys import Keys
def open_browse():
    # 定义H5机型
    mobile_emulation = {
        # "deviceName": "Apple iPhone 3GS"
        # "deviceName": "Apple iPhone 4"
        # "deviceName": "Apple iPhone 5"
        # "deviceName": "Apple iPhone 6"
        # "deviceName": "Apple iPhone 6 Plus"
        # "deviceName": "BlackBerry Z10"
        # "deviceName": "BlackBerry Z30"
        # "deviceName": "Google Nexus 4"
        # "deviceName": "Google Nexus 5"
        # "deviceName": "Google Nexus S"
        # "deviceName": "HTC Evo, Touch HD, Desire HD, Desire"
        # "deviceName": "HTC One X, EVO LTE"
        # "deviceName": "HTC Sensation, Evo 3D"
        # "deviceName": "LG Optimus 2X, Optimus 3D, Optimus Black"
        # "deviceName": "LG Optimus G"
        # "deviceName": "LG Optimus LTE, Optimus 4X HD"
        # "deviceName": "LG Optimus One"
        # "deviceName": "Motorola Defy, Droid, Droid X, Milestone"
        # "deviceName": "Motorola Droid 3, Droid 4, Droid Razr, Atrix 4G, Atrix 2"
        # "deviceName": "Motorola Droid Razr HD"
        # "deviceName": "Nokia C5, C6, C7, N97, N8, X7"
        # "deviceName": "Nokia Lumia 7X0, Lumia 8XX, Lumia 900, N800, N810, N900"
        # "deviceName": "Samsung Galaxy Note 3"
        # "deviceName": "Samsung Galaxy Note II"
        # "deviceName": "Samsung Galaxy Note"
        # "deviceName": "Samsung Galaxy S III, Galaxy Nexus"
        # "deviceName": "Samsung Galaxy S, S II, W"
        # "deviceName": "Samsung Galaxy S4"
        # "deviceName": "Sony Xperia S, Ion"
        # "deviceName": "Sony Xperia Sola, U"
        # "deviceName": "Sony Xperia Z, Z1" #"deviceName": "Amazon Kindle Fire HDX 7″"
        # "deviceName": "Amazon Kindle Fire HDX 8.9″"
        # "deviceName": "Amazon Kindle Fire (First Generation)"
        # "deviceName": "Apple iPad 1 / 2 / iPad Mini"
        # "deviceName": "Apple iPad 3 / 4"
        # "deviceName": "BlackBerry PlayBook"
        # "deviceName": "Google Nexus 10"
        # "deviceName": "Google Nexus 7 2"
        # "deviceName": "Google Nexus 7"
        # "deviceName": "Motorola Xoom, Xyboard"
        # "deviceName": "Samsung Galaxy Tab 7.7, 8.9, 10.1"
        # "deviceName": "Samsung Galaxy Tab"
        # "deviceName": "Notebook with touch"\
        'deviceName': 'iPhone X'
        # Or specify a specific build using the following two arguments
        # "deviceMetrics": { "width": 360, "height": 640, "pixelRatio": 3.0 },
        # "userAgent": "Mozilla/5.0 (Linux; Android 4.2.1; en-us; Nexus 5 Build/JOP40D) AppleWebKit/535.19 (KHTML, like Gecko) Chrome/18.0.1025.166 Mobile Safari/535.19" }
    }
    options = webdriver.ChromeOptions()
    options.add_experimental_option('mobileEmulation', mobile_emulation)
    driver = webdriver.Chrome(chrome_options=options)
    return driver


# 打开并最大化网页窗口
def open_homepage(browse, url):
    # 最大化浏览器
    browse.maximize_window()
    # 打开康美首页地址
    browse.get(url)
def open_xlrd():
    # 打开用例文件，读取对应用例的用户名等数据EE:\PythonProject\mrbtest\kmjs\kmh5
    casefile = xlrd.open_workbook('E:\\PythonProject\\mrbtest\\kmjs\\kmh5\\H5TestData.xls', formatting_info=True)
    # 设置日期格式
    style1 = xlwt.XFStyle()
    style1.num_format_str = 'YYYY-MM-DD HH:MM:SS'
    # 设置单元格背景颜色
    font0 = xlwt.Font()
    font0.name = 'Times New Roman'  # 字体
    font0.colour_index = 2  # 颜色
    font0.bold = True  # 加粗
    style2 = xlwt.XFStyle()
    style2.font = font0
    # 准备向用例文件中写入测试结果
    wb = copy(casefile)
    ws = wb.get_sheet(0)
    # 打开第一张表
    table = casefile.sheets()[0]
    print u"开始执行"
    return ws,table,wb,style1,style2


# 账号密码登陆并判断是否登陆成功
def get_user_login(browse, name, pwd):
    browse.find_element_by_xpath('//*[@id="app"]/div/footer/div/div[4]/div[1]/p').click()
    time.sleep(2)
    browse.find_element_by_xpath('//*[@id="app"]/div/section/div[1]/input').send_keys(name)
    time.sleep(1)
    browse.find_element_by_xpath('//*[@id="app"]/div/section/div[2]/input').send_keys(pwd)
    # 点击登录
    browse.find_element_by_xpath('//*[@id="app"]/div/section/div[3]/button').click()
    time.sleep(1)
    browse.implicitly_wait(3)
    # 判断是否登录成功，如果左上角出现“欢迎来到艾美e族商城”，则判断用户登录成功
    nickname=browse.find_element_by_xpath('//*[@id="app"]/div/div/div[1]/span').text
    print nickname
    if nickname=='飞扬':
        print "用户登录成功"
        return 0
    else:
        print '登陆失败'
        return 1