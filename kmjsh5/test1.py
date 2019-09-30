# -*- coding=utf-8 -*-
# @Time     :2019/9/29 16:07
# @Author   :ZhouChuqi
import kmjsh5
(ws,table,wb,style1,style2)=kmjsh5.open_xlrd()
print ws,table,wb,style1,style2
userName = table.cell(8, 5).value
print userName