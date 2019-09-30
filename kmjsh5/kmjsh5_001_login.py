# -*- coding=utf-8 -*-
# @Time     :2019/9/28 17:25
# @Author   :ZhouChuqi
# *************************************
# *******已注册会员购买商品时提示登录******
# *************************************
# 导入依赖模块
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
print u"****Case--AmezMallH5_001_Login已注册会员购买商品时提示登录--开始运行****"

try:
    # 失败标志
    errorFlag = 0
    # 读取用户名
    userName = table.cell(8, 5).value
    print userName
    # 读取密码
    passWord = table.cell(9, 5).value
    print passWord
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

    # 打开谷歌浏览器
    options = webdriver.ChromeOptions()
    options.add_experimental_option('mobileEmulation', mobile_emulation)
    driver = webdriver.Chrome(chrome_options=options)
    # 最大化浏览器
    driver.maximize_window()
    # 打开康美首页地址
    loginadress = table.cell(3, 5).value
    driver.get(loginadress)
    # 搜索栏输入要搜索的商品
    driver.find_element_by_xpath('//*[@id="app"]/div/footer/div/div[3]/div[1]/p').click()
    SearchGoods = table.cell(4, 5).value
    # 点击搜索栏进入搜索页面
    driver.find_element_by_xpath('//*[@id="input"]').click()
    time.sleep(1)
    # 输入要搜索的商品
    driver.find_element_by_xpath('//*[@id="app"]/div/div[1]/div/div[2]/div/div/div[2]/div/input').send_keys(SearchGoods, Keys.ENTER)
    time.sleep(3)
    # 查看是否出现需要的商品
    ShowGoods = driver.find_element_by_xpath('//*[@id="app"]/div/div[2]/div/ul/li/div/div[2]/p').text
    ReadGoods = table.cell(5, 5).value
    print ReadGoods
    print u"搜索出来的商品为：", ShowGoods
    if (ShowGoods == ReadGoods):
        print u"搜索商品正确！！"
        # 点击搜索出来的商品进入详情页面
        driver.find_element_by_xpath('//*[@id="app"]/div/div[2]/div/ul/li/div/div[1]/img').click()
        time.sleep(2)
        # JS点击立即购买按钮
        driver.find_element_by_xpath('//*[@id="app"]/div/div[7]/button[2]').click()
        # js = 'document.getElementsByClassName("ljgm")[0].click();'
        # driver.execute_script(js)
        time.sleep(2)
        # 选择产品规格：
        driver.find_element_by_xpath('//*[@id="app"]/div/div[9]/div[2]/div/div[2]/div/button[2]').click()
        # countNumber = driver.find_element_by_xpath(
        #     "//*[@id='app']/div/section[8]/section/div[1]/div[2]/div/div[2]/div[2]").text
        # print "countNumber:", countNumber
        # 小米电视4
        # el = driver.find_element_by_xpath("//*[@id='app']/div/section[8]/section/div[2]/div[1]/div/span[1]")
        # driver.execute_script("arguments[0].click();", el)
        # time.sleep(1)
        # # 75英寸
        # el = driver.find_element_by_xpath("//*[@id='app']/div/section[8]/section/div[2]/div[2]/div/span[1]")
        # driver.execute_script("arguments[0].click();", el)
        # time.sleep(1)
        # # 点击确定
        # js = 'document.getElementsByClassName("but_full_square_outer")[0].click();'
        # driver.execute_script(js)
        time.sleep(3)
        driver.find_element_by_xpath('//*[@id="app"]/div/div[9]/div[3]/button').click()
        time.sleep(3)
    # 跳转至登录页面输入用户名密码
    driver.find_element_by_xpath('//*[@id="app"]/div/section/div[1]/input').send_keys(userName)
    driver.find_element_by_xpath('//*[@id="app"]/div/section/div[2]/input').send_keys(passWord)
    # 点击登录
    driver.find_element_by_xpath('//*[@id="app"]/div/section/div[3]/button').click()
    print "未登录用户购买商品提示登录成功"
    time.sleep(3)
    ws.write(3, 7, 'Pass')
    # 如果成功，将错误日志覆盖
    ws.write(3, 10, '')
    errorFlag = 1
    print (u"Case--kmjs-001-Login已注册会员购买商品时提示登录--结果：Pass!")
except Exception as e:
    print(e)
    # 抛出异常
    traceback.format_exc()
    # 写入异常至用例文件中：
    errorInfo = str(traceback.format_exc())
    print "****errorInfo:", errorInfo
    ws.write(3, 10, errorInfo, style2)

finally:
    if (errorFlag == 0):
        print (u"Case--kmjs-001-Login已注册会员购买商品时提示登录--结果：Failed!")
        ws.write(3, 7, 'Failed', style2)
    # 写入执行人员
    ws.write(3, 9, 'zhouchuqi')
    # 写入执行日期
    ws.write(3, 8, datetime.now(), style1)
    # 利用保存时同名覆盖达到修改excel文件的目的,注意未被修改的内容保持不变
    wb.save('E:\\PythonProject\\mrbtest\\kmjs\\kmh5\\H5TestData.xls')
    # 退出浏览器
    driver.quit()
    print u"Case--kmjs-001-Login已注册会员购买商品时提示登录.py运行结束！！！"
