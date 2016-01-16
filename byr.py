# -*- coding: utf-8 -*-

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import time
import re
from xlrd import open_workbook
from xlutils.copy import copy
import os

MAX_COUNT_TIMES = 30
COUNT_INTERVAL = 1800
MINIMIZE = False
AUTO_MAKE_XLS = False

#获取指定分区号下所有版块的版块英文名和在线人数，返回一个列表
def find_sec(secid):
    #正则模块
    pa = re.compile(r'\w+')
    
    browser = webdriver.Firefox() # Get local session of firefox
    
    if MINIMIZE:
        browser.set_window_size(10, 10)#最小化浏览器，没试过这样行不行
    browser.get("http://bbs.byr.cn/#!section/%d "%secid) # Load page
        
    time.sleep(1) # Let the page load
    result = []
    try:
        #获得版面名称和在线人数，存入列表中
        board = browser.find_elements_by_class_name('title_1')
        ol_num = browser.find_elements_by_class_name('title_4')
        
        max_bindex = len(board)
        max_oindex = len(ol_num)
        assert max_bindex == max_oindex,'index not equivalent!'
        
        #版面名称有中英文，用正则过滤只剩英文的，存入列表
        for i in range(1,max_oindex):
            board_en=pa.findall(board[i].text)
            result.append([str(board_en[-1]),int(ol_num[i].text)])
            
        browser.close()
        return result
    except NoSuchElementException:
        assert 0, "can't find element"

#写入excel，xlutils可以写入到已存在的excel中,xlwt只能每次都重写
def write_xls(lis,column,filename):
    rb = open_workbook(filename)
    wb = copy(rb)
    ws = wb.get_sheet(0)
    data_time = time.ctime().split(' ')[3]
    ws.write(0,0,'')#第一行第一列显示空
    ws.write(0,column,data_time)#第一行显示时间
    for i in range(1,len(lis)+1):
        #写入excel中，第0列是板块名称，第1,2,3...列是在线人数
        ws.write(i,0,lis[i-1][0])
        ws.write(i,column,lis[i-1][1])
    wb.save(filename)

def main():
    if AUTO_MAKE_XLS:
        #自动创建xls文件，会提示格式损坏。所以还是手动在当前目录创建一个count.xls文件吧= =
        with open(os.getcwd() + '\\count.xls', 'wb') as f:
            #windows用\\，其他系统用/
            f.write('')
    
    cnt_times=1
    for cnt_times in range(MAX_COUNT_TIMES): 
        #获得分区2（学术科技）下面所有板块在线人数，在终端打印并写入count.xls中
        result_lis = find_sec(2)
        print(result_lis)
        write_xls(result_lis,cnt_times,'count.xls')
    
        cnt_times += 1
        #每隔指定时间统计一次(考虑5s的运行时间)
        time.sleep(COUNT_INTERVAL - 5)
        
if __name__ == '__main__':
    main()
