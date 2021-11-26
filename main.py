import requests
import bs4
import re
import time
import xlwt
import xlrd
import telnetlib
import urllib 
from bs4 import BeautifulSoup
from urllib import request
from urllib.request import urlopen
from my_fake_useragent import UserAgent 
from datetime import datetime
import datetime
import cv2
import os

def time_dif(min,m):     #计算待截取位置与文件起始位置的时间差（截取时间，文件起点）
    s1 = int (min%100)
    s2 = int(m%100)
    m1 = int((min/100)%100)
    m2 = int((m/100)%100)
    h1 = int((min/10000)%100)
    h2 = int((m/10000)%100)
    b = (h1-h2)*3600
    b += (m1-m2)*60
    b += (s1-s2)
    #print(b)
    return b
def getTimeDiff(timeStra,n):        #计算截取时间
    ta = time.strptime(timeStra, "%Y-%m-%d %H:%M")
    y,m,d,H,M = ta[0:5]
    if d != 13 : return
    if H >= 23 or H < 5: return
    if (M >= 43 and n == 18 )or(M >= 56 and n == 5):    
        H = H + 1
        M_1 = (M+n-1)%60
        tmin=datetime.datetime(y,m,d,H,M_1,30)
        min = 30+M_1*100+H*10000+d*1000000+m*100000000+y*10000000000
    else:
        M_1 = M+n-1
        tmin=datetime.datetime(y,m,d,H,M,30)
        min = 30+M_1*100+H*10000+d*1000000+m*100000000+y*10000000000
    if (M >= 42 and n == 18 )or(M >= 55 and n == 5):    
        H = H + 1
        M_1 = (M+n)%60
        tmax=datetime.datetime(y,m,d,H,M_1,30)
        max = 30+M_1*100+H*10000+d*1000000+m*100000000+y*10000000000
    else :
        M_1=M+n
        tmax=datetime.datetime(y,m,d,H,M,30)
        max = 30+M_1*100+H*10000+d*1000000+m*100000000+y*10000000000
    #print(tmin)
    #print(tmax)
    print(min)
    print(max)
    return (min,max)

def get_xlsx():
    xlsx_path = 'home\henrryzh\\ncc\\reid\\2021-09-13.xlsx'
    #xlsx_path = 'D:\人脸识别\南六楼抓拍刷卡list\刷卡记录\\1-3月南六楼记录.xlsx'
    #xlsx_path = 'D:\人脸识别\南六楼抓拍刷卡list\刷卡记录\\0401_0430.xlsx'
    data_xsls = xlrd.open_workbook(xlsx_path)
    sheet_name = data_xsls.sheets()[0]
    count_nrows = sheet_name.nrows  #获取总行数
    datalist = []
    
    for i in range(2,count_nrows):
        a=sheet_name.cell(i,4).value  #根据行数来取对应列的值，并添加到字典中
        datalist.append(a)
        b=sheet_name.cell(i,3).value  #打卡位置
        c=sheet_name.cell(i,1).value  #姓名
        if (c == ''): continue
        '''
    a=sheet_name.cell(5522,4).value  #根据行数来取对应列的值，并添加到字典中
    datalist.append(a)
    b=sheet_name.cell(5522,3).value  #打卡位置
    c=sheet_name.cell(5522,1).value  #姓名
    if (c == ''): pass#return 
    '''
        if (b == '后门出口' or '后门进口'): #n时间差cam摄像头号
            n = 18
            cam = 29
        else: 
            n = 5
            cam = 22
        (min,max)=getTimeDiff(a,n)
        if(cam == 29):
            d = os.listdir('home/henrryzh/ncc/reid/20210913_29')
            for video in d:
                M1 = int(video.split("_")[3])
                M2 = int(video.split("_")[4])
                if(M1<=min and M2>=max): break
            print(video)
            cap = cv2.VideoCapture(os.path.join('home/henrryzh/ncc/reid/20210913_29',video))    
        else:
            d = os.listdir('home/henrryzh/ncc/reid/20210913_22')
            for video in d:
                M1 = int(video.split("_")[3])
                M2 = int(video.split("_")[4])
                if(M1<=min and M2>=max): break
            print(video)
            cap = cv2.VideoCapture(os.path.join('home/henrryzh/ncc/reid/20210913_22',video))
            
        c = 1
        timeRate = 2  # 截取视频帧的时间间隔（这里是每隔2秒截取一帧）
        a = 1
        while(True):
            ret, frame = cap.read()
            FPS = cap.get(5)
            
            if ret:
                frameRate = int(FPS) * timeRate  # 因为cap.get(5)获取的帧数不是整数，所以需要取整一下（向下取整用int，四舍五入用round，向上取整需要用math模块的ceil()方法）
                if(c >= (time_dif(min,M1))*int(FPS) and c <= (time_dif(max,M1))*int(FPS)):
                    
                    if(c % frameRate == 0):
                        print("开始截取视频第：" + str(a) + " 帧")
                        # 这里就可以做一些操作了：显示截取的帧图片、保存截取帧到本地
                        #print("D:/人脸识别/jpg/" + str(a) + '.jpg')
                        cv2.imwrite("D:/json/" + str(a) + '.jpg', frame)  # 这里是将截取的图像保存在本地
                        a += 1
                c += 1
                cv2.waitKey(0)
            else:
                print("所有帧都已经保存完成")
                break

    #print(datalist[0])
    #print(count_nrows)
    # return
if __name__ == "__main__":
    get_xlsx()



