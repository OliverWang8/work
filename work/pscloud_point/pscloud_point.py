# -*- coding:utf-8 -*-
from  selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import configparser
import time, unittest, os
chrome_options = Options()
chrome_options.add_argument('---headless')
wd = webdriver.Chrome(chrome_options=chrome_options)
def getConfig(section, key):
    config = configparser.ConfigParser()
    path = os.path.split(os.path.realpath(__file__))[0] + '/config.conf'
    config.read(path)
    return config.get(section, key)
url = getConfig("pscloud","url")
mysql = getConfig("pscloud","mysql")
name = getConfig("pscloud","name")
password = getConfig("pscloud","password")
is_choise = getConfig('pscloud','is_choise')
is_test = getConfig('pscloud','is_test')
wd.get(url)
if is_test == 'yes':
    wd.find_element_by_link_text(mysql).click()
wd.find_element_by_id("login").send_keys(name)
wd.find_element_by_id("password").send_keys(password)
wd.find_element_by_class_name("btn-primary").click()
time.sleep(2)
wd.find_element_by_link_text("总账").click()
time.sleep(1)

search_list = ['总分类账','明细分类账','核算维度明细账','多栏式明细账','科目余额表','核算维度余额表']
with open('pscloud_point.txt', 'w') as f:
    if is_choise == 'yes':
        f.write('勾选未记账选项响应时间如下：')
        f.write('\r\n')
        print('勾选未记账选项响应时间如下：')
    else:
        f.write('未勾选未记账选项响应时间如下：')
        f.write('\r\n')
        print('未勾选未记账选项响应时间如下：')
    for a in search_list:
        time.sleep(1)
        if a == '科目余额表' or a == '核算维度余额表':
            wd.find_element_by_link_text("财务报表").click()
        else:
            wd.find_element_by_link_text("账簿").click()
        time.sleep(1)
        wd.find_element_by_link_text(a).click()
        time.sleep(2)
        account_name = wd.find_element_by_class_name('o_control_panel')
        if is_choise == 'yes':
            if a == '总分类账' or a == '科目余额表' :
                wd.find_element_by_name('is_posted').click()
            elif a == '多栏式明细账':
                wd.find_element_by_name('is_contains_unaccounted').click()
            elif a == '明细分类账' or a == '核算维度明细账' :
                wd.find_element_by_name('is_unaccounted_voucher').click()
            else:
                wd.find_element_by_name('is_posted').click()
                time.sleep(1)
                choise_weidu = wd.find_elements_by_class_name('ui-autocomplete-input')[2]
                choise_weidu.click()
                choise_weidu.send_keys('往来单位')
                time.sleep(1)
                choise_weidu.send_keys(Keys.TAB)
                choise_weidu = wd.find_elements_by_class_name('ui-autocomplete-input')[2]
                choise_weidu.send_keys('核算产品')
                time.sleep(1)
                choise_weidu.send_keys(Keys.TAB)
                choise_weidu = wd.find_elements_by_class_name('ui-autocomplete-input')[2]
                choise_weidu.send_keys('现金流量')
                time.sleep(1)
                choise_weidu.send_keys(Keys.TAB)
                choise_weidu = wd.find_elements_by_class_name('ui-autocomplete-input')[2]
                choise_weidu.send_keys('部门')
                time.sleep(1)
                choise_weidu.send_keys(Keys.TAB)
                choise_weidu = wd.find_elements_by_class_name('ui-autocomplete-input')[2]
                choise_weidu.send_keys('个人')
                time.sleep(1)
                choise_weidu.send_keys(Keys.TAB)
        elif a == '核算维度余额表':
            choise_weidu = wd.find_elements_by_class_name('ui-autocomplete-input')[2]
            choise_weidu.click()
            choise_weidu.send_keys('往来单位')
            time.sleep(1)
            choise_weidu.send_keys(Keys.TAB)
            choise_weidu = wd.find_elements_by_class_name('ui-autocomplete-input')[2]
            choise_weidu.send_keys('核算产品')
            time.sleep(1)
            choise_weidu.send_keys(Keys.TAB)
            choise_weidu = wd.find_elements_by_class_name('ui-autocomplete-input')[2]
            choise_weidu.send_keys('现金流量')
            time.sleep(1)
            choise_weidu.send_keys(Keys.TAB)
            choise_weidu = wd.find_elements_by_class_name('ui-autocomplete-input')[2]
            choise_weidu.send_keys('部门')
            time.sleep(1)
            choise_weidu.send_keys(Keys.TAB)
            choise_weidu = wd.find_elements_by_class_name('ui-autocomplete-input')[2]
            choise_weidu.send_keys('个人')
            time.sleep(1)
            choise_weidu.send_keys(Keys.TAB)
        start_time = int(round(time.time() * 1000))
        wd.find_element_by_name('search_psreport').click()
        change_style = wd.find_element_by_class_name('o_loading').get_attribute('style')
        success = True
        while success:
            #print('----------')
            change_style = wd.find_element_by_class_name('o_loading').get_attribute('style')
            if ('none' in change_style):
                success = False
        end_time = int(round(time.time() * 1000))
        response_time = float((end_time - start_time)/1000)
        print(a+'查询响应时间：'+str(response_time))
        f.write(a+'查询响应时间：'+str(response_time))
        f.write('\r\n')
time.sleep(5)
wd.quit()