# -*- coding: UTF-8 -*-
#opexcel.py
import numpy as np
import xlwings as xw

wb = xw.Book(r'D:\code\python\test2.xlsm')
sht2 = wb.sheets['Sheet2']
sht3 = wb.sheets['Sheet3']
row2=sht2.range('A1').end('down').row
row3=sht3.range('A1').end('down').row

# 整体设计如下：
# sheet1是个button
# sheet2是昨天的数据
# sheet3是今天的数据
# sheet4是历史数据生成的表格

def DTSWORK():
    #打开昨天和今天的数据
    flag2=0
    flag3=0
    find = 0
    #开始遍历两边的数据并进行处理，从今天的数据开始处理
    while (flag3 < row3):
        flag2 = 0
        find = 0
        while (flag2 < row2):
            #如果今天和昨天有相同的事务，说明昨天没处理完，把昨天的进展拷贝过来，今天继续处理
            if sht3.range('A1').offset(flag3,0).options(numbers=int).value == sht2.range('A1').offset(flag2,0).options(numbers=int).value:
                sht3.range('A1').offset(flag3,3).value = sht2.range('A1').offset(flag2,3).value
                find = 1
                break                
            else:  
                flag2 += 1
        if (find == 0):    #如果没有昨天的数据，说明是今天新增的事务，标个颜色，开搞
            sht3.range('A1').offset(flag3,0).color = (255,100,100)
        flag3 += 1    
    
    #保存文件，关闭文件，是个好习惯
    wb.save()
    #wb.close()
    return;

DTSWORK();

#可能会实现的功能:统计今天剩余的DI

#遗留问题:
#1、python里面怎么实现类似C语言里面的for循环？for range
