# -*- coding:utf-8 -*-
import xlrd,time,xlwt,csv
from xlutils import copy#copy已存在的数据，实现写入功能
from selenium import webdriver
from selenium.webdriver.common.keys import Keys#模拟回车操作
from xlutils.copy import copy as xl_copy
csvName = 'C:\\Users\\wangshuai04\\Downloads\\Jira 2019-08-03T15_07_15+0800.csv'
sheetName = '20190806'
workbookPath = 'F:\\project\\account_status\\PSCloud0830BUG优先级列表.xls'
def csv_to_xlsx():
    with open(csvName, 'r', encoding='utf-8') as f:
        read = csv.reader(f)
        rb = xlrd.open_workbook(workbookPath)
        if sheetName in rb.sheet_names():
            print('该Sheet名称已存在，请换名字！')
            return False
        else:
            print(rb.sheet_names())
            wb = xl_copy(rb)
            sheet = wb.add_sheet(sheetName)  # 创建一个sheet表格
            l = 0
            for line in read:
                r = 0
                for i in line:
                    sheet.write(l, r, i)  # 一个一个将单元格数据写入
                    r = r + 1
                l = l + 1
            print('CSV转换Excel成功！')
            wb.save(workbookPath)  # 保存Excel
            return True
def stateBack():
    userName = 'wangshuai04@inspur.com'#jira账户
    password = 'pllilu1314!@'#jira密码
    option = webdriver.ChromeOptions()
    option.add_argument('headless')
    wd = webdriver.Chrome(chrome_options=option)
    wd.get('http://pm.dev.mypscloud.com')
    wd.find_element_by_name('os_username').send_keys(userName)
    wd.find_element_by_name('os_password').send_keys(password)
    wd.find_element_by_name('login').click()
    time.sleep(2)
    workbook = xlrd.open_workbook(workbookPath)
    newWorkbook = copy.copy(workbook)#复制已有的文件
    sheetNumber = len(workbook.sheets())
    for j in range(0,sheetNumber-1):
        sheet = newWorkbook.get_sheet(j)#读取第x个sheet页
        #sheet1 = workbook.sheet_by_name(sheetName)#根据sheet页名称读数据
        sheet1 = workbook.sheet_by_index(j)
        rows = sheet1.nrows#行数
        cols = sheet1.ncols#列数
        sheet.write(0,cols-1,'状态')
        for i in range(1,rows):
            a_key = sheet1.cell_value(i, 0)#读取jira问题编号
            a = wd.find_element_by_class_name('search')
            a.send_keys(a_key)
            a.send_keys(Keys.ENTER)
            time.sleep(1)
            b = wd.find_element_by_id('resolution-val')#读取问题解决状态
            time.sleep(1)
            b_value = b.text
            sheet.write(i,cols-1,b_value)
            print(a_key+'处理结果为：'+b_value)
    newWorkbook.save(workbookPath)#保存新的文件
    time.sleep(2)
    wd.quit()
    print('回写状态完成！')
if __name__ == '__main__':
    if csv_to_xlsx() == True:
        stateBack()