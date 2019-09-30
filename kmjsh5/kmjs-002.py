# -*- coding=utf-8 -*-
# @Time     :2019/9/29 10:20
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
# 打开用例文件，读取对应用例的用户名等数据
import kmjsh5
#定义数据表格读写格式
(ws,table,wb,style1,style2)=kmjsh5.open_xlrd()
print ws,table,wb,style1,style2
try:
    # 失败标志
    errorFlag = 0
    # 读取用户名
    userName = table.cell(8, 5).value
    print userName
    # 读取密码
    passWord = table.cell(9, 5).value
    print passWord
    loginadress = table.cell(3, 5).value
    # 定义H5机型1
    driver=kmjsh5.open_browse()
    # 打开谷歌浏览器
    kmjsh5.open_homepage(driver, loginadress)
    # 点击个人中心跳转登陆
    code=kmjsh5.get_user_login(driver,userName,passWord)
    print code
    time.sleep(2)
    ws.write(3, 7, u'脚本执行成功')
    # 如果成功，将错误日志覆盖
    if code==0:
        ws.write(3, 10, u'用户登陆成功')
    else: ws.write(3, 10, u'用户登陆失败')

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

