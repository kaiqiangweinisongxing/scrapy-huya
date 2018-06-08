#!/usr/bin/env python 
# -*- coding: utf-8 -*-
import requests
import re
from lxml import etree
import os
import json  
#from operator import itemgetter, attrgetter
import xlwt     #操作excel
import datetime
import time

#http://www.huya.com/
#http://www.huya.com/g/wzry
#封面：<img class="pic" data-original="//screenshot.msstatic.com/yysnapshot/1710dd04a5a7aa7adfca297029740b9d328f5358f43e?imageview/4/0/w/338/h/190/blur/1" src
#标题：   <a href="http://www.huya.com/guying" class="title new-clickstat"report='{"eid":"click/position","position":"wzry/0/1/1","game_id":"2336","ayyuid":"1757672727"}' title="今天有点小帅" ta

    #进入专区enterTheZone
    #获取直播间信息
    #保存图片（以直播间命名）
    #输出信息（json）
    #输出txt文件，格式：
                    #主播名        #观看人数（高到低排序）        #房间地址
#单线程：下载图片26s  .优化后多线程，9.45s

def enterTheZone():
    print("For example, http://www.huya.com/g/wzry, zone is 'wzry'. ")
   # label = input("So please input a zone:")
    label = 'wzry'
    #容错处理
    while(True):
        url = "http://www.huya.com/g/{}".format(label.lower())
        #获取gid（gameid）
        #responce = requests.get(url)
        #selector = etree.HTML(responce.text)
        #text = selector.xpath('//script[@data-fixed="true"][contains(text(),"var GID")]/text()')
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
    

def getInformation(gid):
    '''获得专区所有房间信息 # list->dict 
    '''
    '''
    responce = requests.get(url)
    selector = etree.HTML(responce.text)
    roomList = []
    #直接找页面的html代码
    li = selector.xpath('//li[@class="game-live-item"]')
    for x in li:
        #房间编号
        roomUrl = x.xpath('./a[@class="video-info new-clickstat"]/@href')
        #房间名
        roomName = x.xpath('.//a[@class="title new-clickstat"]/text()')  
        #主播名
        name = x.xpath('.//span[@class="avatar fl"]/i/text()') 
        #观看人数
        number = x.xpath('.//span[@class="num"]/i[@class="js-num"]/text()')
        #封面url
        imgUrl = 'http:' + x.xpath('.//img[@class="pic"]/@data-original')[0]
        dict =  {'roomUrl':roomUrl,'roomName': roomName,'name':name,'number':number,'imgUrl':imgUrl}
        roomList.append(dict)
    return roomList
    '''
    #network中找XHR信息   http://www.huya.com/cache.php?m=LiveList&do=getLiveListByPage&gameId=2336&tagAll=0&page=1   在http://www.bejson.com/ 校验查看      #gid是王者荣耀的id,page
    XHR =  'http://www.huya.com/cache.php?m=LiveList&do=getLiveListByPage&gameId={}&tagAll=0&page={}'
    page = 0
    roomList = []
    while(page<1):
        page += 1
        print("page:",page)
        response = requests.get(XHR.format(gid,page))
        ob_json = json.loads(response.text) #转换为字典
        datas_list = ob_json.get('data').get('datas')
        if( datas_list == []):
            return roomList
        else:
            for d in datas_list:
                dict = {}       #地址传递，别写在for外面
                #观看人数
                #number = dict["totalCount"]
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

                #获取订阅人数和开播时间
                regx = '//script[@data-fixed="true"][contains(text(),"var TT_ROOM_DATA")]/text()'
                url = 'http://www.huya.com/{}'.format(d["privateHost"])
                #text = fc.send(url)                                      #协程
                text = findFromHtml(regx,url)
                startTime = sliceUp(text[0],'"startTime":"','","')
                dict["startTime"] = timestamp2string(int(startTime))
                dict["fans"] =  sliceUp(text[0],'"fans":',',"')
                roomList.append(dict)
    return roomList  

'''
#协程
def findFromHtmlCoroutine():
        regx = '//script[@data-fixed="true"][contains(text(),"var TT_ROOM_DATA")]/text()'
        data = ''
        while(True):
            data = yield data
            responce = requests.get(data)
            selector = etree.HTML(responce.text)
            data = selector.xpath(regx)
'''          

def timestamp2string(timeStamp):  
        '''时间戳转换为字符串日期时间
        timeStamp:int 类型
        返回： str类型
        '''
        try:  
            d = datetime.datetime.fromtimestamp(timeStamp)  #2015-08-28 16:43:37
            return d.strftime("%Y-%m-%d %H:%M:%S")    # 2015-08-28 16:43:37.283000 
        except Exception as e:  
            print (e) 
            return ''  
  
'''
def saveInformation(roomList,zoneName):
    #保存信息到当前目录下
    
    count = 1
    imgPath = os.getcwd() + '\\' + zoneName 
    if not os.path.exists(imgPath):
       os.mkdir(imgPath)
    jsonPath = os.getcwd() + '\\' + zoneName + '.json' 
    f1 = open(jsonPath,'a', encoding='utf-8')
    for roomDict in roomList:
        thread_lock.acquire()  # 上锁
        t = threading.Thread(target=downloadImg, args=(roomDict, imgPath))
        t.start()
        print(count )       #显示进度
        count += 1
        #写入文本，记录所有信息
        for key,value in roomDict.items():
            f1.write('{}:{}'.format(key,value))
            f1.write('\t\t')
        f1.write('\n')
    f1.close()

def downloadImg(roomDict,imgPath):
        content = requests.get(roomDict['imgUrl']).content
        #保存图片
        with open(imgPath +'\\' + fileNameFilter(roomDict["name"]) + '.jpg','wb') as f:
           f.write(content)
        thread_lock.release()  # 解锁

    
def stringToInt(string):
    #str转int
    try:
        if(string.find('万') != -1 ): 
            return int(string[0:string.find('万')]) * 10000
        else:       #整数
            return int (string)
    except ValueError as e: #有浮点数
        return int(float(string[0:string.find('万')]) * 10000)


def fileNameFilter(fileName):
    #过滤文件命名规则的禁用字符
    
    ban = ('\\','/',':','*','?','"','<','>','|')
    for x in ban:
        fileName = fileName.replace(x,' ')
    return fileName
'''

def saveToExcel(roomList,zoneName):
    '''保存到excel文件
    '''
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
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet('热门排行榜')
    worksheet.write(0,0,'直播间',sty1)
    worksheet.write(0,1,label = '主播',style = sty1)
    worksheet.write(0,2,label = '描述',style = sty1)
    worksheet.write(0,3,label = '观众',style = sty1)
    worksheet.write(0,4,label = '粉丝',style = sty1)
    worksheet.write(0,5,label = '开播时间',style = sty1) 
    worksheet.write(0,6,label = '传送门',style = sty1)
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
    except PermissionError as e:
        print(e)
        os.system('pause') 


if __name__ == '__main__':
    zone = enterTheZone()
    old = time.time()
    roomList = getInformation(zone[0])
    saveToExcel(roomList,zone[1])
    print('耗时:',time.time() - old)