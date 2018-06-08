#!/usr/bin/env python 
# -*- coding: utf-8 -*-
import requests
import re
from lxml import etree
import os
import json  
import xlwt     #操作excel
import datetime
import time
import threading  # 多线程

# 设置最大线程锁
thread_lock = threading.BoundedSemaphore(value = 30)

#全局变量
totalCount = 0      #房间总数
allRoomList = []
runTime = 0

def enterTheZone():
    print("For example, http://www.huya.com/g/wzry, zone is 'wzry'. ")
    label = input("So please input a zone:")
    #容错处理
    while(True):
        url = "http://www.huya.com/g/{}".format(label.lower())
        regx = '//script[@data-fixed="true"][contains(text(),"var GID")]/text()'
        text = findFromHtml(regx,url)
        if(text == []):
            label = input('please input a correct zone:')
        else:
            pattern  = re.compile("\d+")
            gid = pattern.findall(text[0])
            return gid[0],label
 
def findFromHtml(regx,url):
        responce = requests.get(url)
        selector = etree.HTML(responce.text)
        return selector.xpath(regx)


#切取中间字段，例如asbvbasid ，获取bas
def sliceUp(string,start,end):
    if((string.find(start)== -1) or (string.find(end)== -1)):        #存在性
        return ''
    length = len(start)
    start = string.index(start)     
    end = string[start:].index(end) + start   
    return string[start + length:end]
    

def getDatas_list(gid):
    '''获得专区所有房间信息 # list->dict 
    '''
    global totalCount
    XHR =  'http://www.huya.com/cache.php?m=LiveList&do=getLiveListByPage&gameId={}&tagAll=0&page={}'
    page = 0
    roomList = []
    urlList = []            #存每页url
    datas_list = []         #所有页所有房间的datas list
    while(True):
        page += 1
        print("page:",page)
        response = requests.get(XHR.format(gid,page))
        ob_json = json.loads(response.text) #转换为字典
        datas_list_page = ob_json.get('data').get('datas')
        if(datas_list_page == []):
            totalCount = len(datas_list)
            return datas_list
        else:
            datas_list.extend(datas_list_page)

        #测试
        '''
        if(page == 1):
            totalCount = len(datas_list)
            return datas_list
        '''
       
def readDatas(datas_list):
    '''利用多线程，读每个房间的数据
    '''
    for d in datas_list:
        thread_lock.acquire()  # 上锁
        t = threading.Thread(target=readInThread, args=(d,0))
        t.start()

def readInThread(d,i):          #i是凑数的，不然出bug
        global allRoomList
        dict = {}
        #观看人数
        dict["number"] = d.get("totalCount")
        #房间名
        dict["roomName"] = d["roomName"]
        #主播
        dict["name"] = d["nick"]
        #房间描述
        dict["introduction"] = d["introduction"]
        #房间封面
        dict["imgUrl"] = d["screenshot"]
        #房间链接
        dict["roomUrl"] = 'http://www.huya.com/{}'.format( d["privateHost"])
        regx = '//script[@data-fixed="true"][contains(text(),"var TT_ROOM_DATA")]/text()'
        url = 'http://www.huya.com/{}'.format(d["privateHost"])  
        text = findFromHtml(regx,url) 
        startTime = sliceUp(text[0],'"startTime":',',').replace('"','')
        #print(startTime)
        if(startTime.strip() == ''):
            print("url:"+ dict["roomUrl"])
        dict["startTime"] = timestamp2string(int(startTime))
        dict["fans"] =  sliceUp(text[0],'"fans":',',"')
        allRoomList.append(dict)
        thread_lock.release()  # 解锁
        judge()

def judge():
    '''判断是否全部数据都读取完毕
    '''
    global totalCount
    global allRoomList
   # print(totalCount)
    totalCount = totalCount - 1
    if(totalCount == 0):
        saveToExcel(allRoomList,"wzry")

def timestamp2string(timeStamp):  
        '''时间戳转换为字符串日期时间
        timeStamp:int 类型
        返回： str类型
        '''
        try:  
            d = datetime.datetime.fromtimestamp(timeStamp)  # 2015-08-28 16:43:37
            return d.strftime("%Y-%m-%d %H:%M:%S")          # 2015-08-28 16:43:37.283000 
        except Exception as e:  
            print (e) 
            return ''  
  
def saveToExcel(roomList,zoneName):
    '''保存到excel文件
    '''
    global runTime
    #根据观看人数排序
    if(roomList == []):
        return
    for dict in roomList:
        #dict["number"] = stringToInt(dict["number"][0])
        dict["number"] = int(dict["number"])
    roomList.sort(key=lambda d:d["number"],reverse = True)
    #生成表格
    font = xlwt.Font()
    font.bold = True
    font.height = 20 * 14
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER 
    alignment.vert = xlwt.Alignment.VERT_CENTER
    sty1 = xlwt.XFStyle()
    sty1.font = font
    sty1.alignment = alignment 
    sty2 = xlwt.XFStyle()
    sty2.alignment = alignment
    row = 1
    column = 0
    workbook = xlwt.Workbook(encoding = 'ascii')        #创建一个Workbook对象
    worksheet = workbook.add_sheet('热门排行榜')
    worksheet.write(0,0,'直播间',sty1)
    worksheet.write(0,1,label = '主播',style = sty1)
    worksheet.write(0,2,label = '描述',style = sty1)
    worksheet.write(0,3,label = '观众',style = sty1)
    worksheet.write(0,4,label = '粉丝',style = sty1)
    worksheet.write(0,5,label = '开播时间',style = sty1) 
    worksheet.write(0,6,label = '传送门',style = sty1)

    #设置行高、列宽
    worksheet.col(0).width = 256 * 46       #宽度45
    worksheet.col(1).width = 256 * 20 
    worksheet.col(2).width = 256 * 40
    worksheet.col(3).width = 256 * 10
    worksheet.col(4).width = 256 * 10
    worksheet.col(5).width = 256 * 20 
    worksheet.col(6).width = 256 * 28


    for dict in roomList:
        worksheet.write(row,column, dict["roomName"],sty2 )
        worksheet.write(row,column + 1,dict["name"] ,sty2)
        worksheet.write(row,column + 2,dict["introduction"],sty2)
        worksheet.write(row,column + 3,dict["number"],sty2)
        worksheet.write(row,column + 4,dict["fans"],sty2)
        worksheet.write(row,column + 5,dict["startTime"],sty2)
        worksheet.write(row,column + 6,xlwt.Formula('HYPERLINK("{}";"{}")'.format(dict["roomUrl"],dict["roomUrl"])))    #超链接
        row += 1
    try:
        workbook.save('{}.xls'.format(zoneName))
        print('耗时:',time.time() - runTime)
        os.system('pause') 
    except PermissionError as e:
        print(e)
        os.system('pause') 


if __name__ == '__main__':
    zone = enterTheZone()
    runTime = time.time()
    datas_list = getDatas_list(zone[0])
    readDatas(datas_list)