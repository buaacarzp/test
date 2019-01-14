# -*- coding: utf-8 -*-
"""
Created on Sun Jan 13 18:17:16 2019

@author: peng.zhou
"""

import xlwt                #引入xlwt数据库用来将数据写入excel文档中
import xlrd                #引入xlrd数据库用来从excel中读取数据
import random              #引入random数据库给出随机数

excel=xlrd.open_workbook('C:\Users\My\Desktop\good.xlsx')              #从一个已有学号与姓名的excel表格中提取出对应的学号和姓名

sheet=excel.sheets()[0]
wb=xlwt.Workbook()                                    #创立一个新的excel表格
ws1=wb.add_sheet('1班成绩单')                          #第一页命名为1班成绩单
ws2=wb.add_sheet('2班成绩单')                          #第一页命名为2班成绩单


a1=[]                                                 #表1中的学号列
a2=[]                                                 #表2中的学号列
b1=[]                                                 #表1中的姓名列
b2=[]                                                 #表2中的姓名列

for i in range (1,14):                                #表1.xlsx表中提取出1班的姓名与学号
      a1.append(sheet.row_values(i,1,2))
      b1.append(sheet.row_values(i,2,3))

      
for j in range (14,33):                               #在表1.xlsx表中提取出1班的姓名与学号
      a2.append(sheet.row_values(j,1,2))
      b2.append(sheet.row_values(j,2,3))
      
for n in range(13):                                   #将1班学号与姓名写入新建的表格中，并写在第1页.1班人数为13人
      ws1.write(n,0,a1[n][0])
      ws1.write(n,1,b1[n][0])
for m in range(19):                                   #将2班学号与姓名写入新建的表格中，并写在第2页.2班人数为19人
      ws2.write(m,0,a2[m][0])
      ws2.write(m,1,b2[m][0])

for q in range(13):                                   #对1班所有人的成绩进行随机抽取数据
      ran=random.randint(60,91)                       #分数为60-90之间
      if ran<=70:
            ws1.write(q,2,'及格({0})'.format(ran))     #以下表示在各分数段的等级
      if ran>70 and ran<=80:
            ws1.write(q,2,'中等({0})'.format(ran))
      if ran>80 and ran<=90:
            ws1.write(q,2,'良好({0})'.format(ran))

for d in range(19):                                    #对2班所有人的成绩进行随机抽取数据
      ran=random.randint(60,91)
      if ran<=70:
            ws2.write(d,2,'及格({0})'.format(ran))
      if ran>70 and ran<=80:
            ws2.write(d,2,'中等({0})'.format(ran))
      if ran>80 and ran<=90:
            ws2.write(d,2,'良好({0})'.format(ran))

      

wb.save('15资环1，2班地理信息系统实习成绩.xls')            #将新建的表格保存为'15资环1，2班地理信息系统实习成绩.xls'文件