# -*- coding:utf-8 -*-
import csv
import xlwt,xlrd
from xlutils.copy import copy as xl_copy
csvName = 'C:\\Users\\wangshuai04\\Downloads\\Jira 2019-08-03T15_07_15+0800.csv'
sheetName = '20190803'
workbookPath = 'PSCloud0830BUG优先级列表.xls'
def csv_to_xlsx():
    with open(csvName, 'r', encoding='utf-8') as f:
        read = csv.reader(f)
        rb = xlrd.open_workbook(workbookPath)
        wb = xl_copy(rb)
        sheet = wb.add_sheet(sheetName)  # 创建一个sheet表格
        l = 0
        for line in read:
            #print(line)
            r = 0
            for i in line:
                # print(i)
                sheet.write(l, r, i)  # 一个一个将单元格数据写入
                r = r + 1
            l = l + 1
        wb.save('PSCloud0830BUG优先级列表.xls')  # 保存Excel

if __name__ == '__main__':
    csv_to_xlsx()
