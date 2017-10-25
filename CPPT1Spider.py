# -*- coding:utf8 -*-

import requests
import os
from lxml import etree
from requests.exceptions import ReadTimeout,ConnectionError,RequestException

LOCALPATH = "/Users/Sophia/Codes/python/spider/pptDownLoad/"
HOMEPAGE  = 'http://www.1ppt.com'

'''
简单封装了ppt1.com单站爬虫的分步功能，大体上分三个步骤：
1、根据首页，获取板块分页目录  GetIndexPage()；
2、遍历板块中的所有文档内容   GetContentByURL；
3、获取文档信息             GetDocInfo

扩展功能1:提供文件下载功能   DownLoadFile
'''
class CPPT1Spider():
    def __init__(self , homePage):
        self.initWithHomePage(homePage)
    
    def initWithHomePage(self , oHomePage):
        print ""
        html = requests.get(oHomePage)
        html.encoding = 'gb2312'
        selector = etree.HTML(html.text)

        self.m_PageListDes = selector.xpath('//dd[@class="ikejian_col_nav"]/ul/li/h4/a/@title')
        self.m_PageListURL = selector.xpath('//dd[@class="ikejian_col_nav"]/ul/li/h4/a/@href')

    #根据首页，获取板块目录。在ppt1网站中，板块目录就是初【出版社->年级教材】
    def GetIndexPage(self):   
        oCounts = 1             
        for value1,value2 in zip(self.m_PageListDes ,self.m_PageListURL):
            oListPageURL = HOMEPAGE + value2
            createPath = LOCALPATH + value1
            print "[", oCounts, "]", value1, oListPageURL
            oCounts = oCounts + 1        
            self.GetContentByURL(oListPageURL , createPath) 
            #if not os.path.exists(createPath):
                #os.mkdir(createPath)
                #PS_getContentPage.GetContentByURL(oListPageURL , createPath)
            #print createPath

    #根据板块首页，遍历所有详情页
    def GetContentByURL(self , oListPageURL , oLocalPath):

        homePage = "http://www.1ppt.com"
        html = requests.get(oListPageURL)
        html.encoding = 'gb2312'
        selector = etree.HTML(html.text)

        # listPage 的获取办法
        content3 = selector.xpath('//ul[@class="arclist"]/li/a/@href')
        for value in content3:
            listPageUrl = homePage + value
            self.GetDocInfo(listPageUrl)
            #oRet = PS_downloadFile.downLoadByURL(listPageUrl , oLocalPath)
            
        #have nextpage 通过“下一页”的关键字查找
        content4 = selector.xpath(u"//a[contains(text(), '下一页')]")
        for value in content4:
            strlistUrl = value.get('href')
            #herf绝对路径的方法没找到，就用字符串拼吧
            op = oListPageURL.rfind('.')
            if (op > 0):
                op1 = oListPageURL.rfind('/') + 1
                strRet = oListPageURL[:op1]
                strRet = strRet.lower()
                oListPageURL = strRet

            nextPage = oListPageURL + strlistUrl
            print "next Page -----------------------"
            print nextPage
            self.GetContentByURL(nextPage , oLocalPath)
    
    #根据详情页地址，获取文档信息。ppt1中的docInfo有很多，我只取了其中三项。
    def GetDocInfo(self , oUrl):
        html = requests.get(oUrl)
        html.encoding = 'gb2312'        
        selector = etree.HTML(html.text)
        zipUrl  = selector.xpath('//ul[@class="downurllist"]/li/a/@href')
        if not zipUrl:
            return  False

        print "课件详情页:" ,oUrl     
        strZip = str(zipUrl[0])
        print "课件地址:" ,strZip
        #下载文件
        #oRet = self.DownLoadFile(strZip)

        #频道地址
        #docInfoList  = selector.xpath('//div[@class="info_left"]/ul/li[1]/a/@href')

        #课件名称
        docInfoList  = selector.xpath('//div[@class="ppt_info clearfix"]/h1/text()')
        if docInfoList:
            print "课件名称:" , docInfoList[0]
        
        print ""
    
    def DownLoadFile(self, oFileURL):

        #模拟报文头
        oHeader = { "Accept":"text/html,application/xhtml+xml,application/xml;",
                "Accept-Encoding":"gzip",
                "Accept-Language":"zh-CN,zh;q=0.8",
                "Referer":"http://www.1ppt.com/",
                "User-Agent":"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.90 Safari/537.36"
                }

        PDFName = self.GetFileNameByURL(oFileURL)
        localPDF = LOCALPATH + PDFName

        try:
            response = requests.get(oFileURL,headers = oHeader ,timeout=10)
            print(response.status_code)
            print localPDF

            if not os.path.exists(localPDF):
                oFile = open(localPDF, "wb")
                for chunk in response.iter_content(chunk_size=512):
                    if chunk:
                        oFile.write(chunk)
            else:
                print "had download file:" ,localPDF

            return True
        except ReadTimeout:
            print("timeout")
        except ConnectionError:
            print("connection Error")
        except RequestException:
            print("error")

        return False

    def GetFileNameByURL(self, oURL):
        op = oURL.rfind('/')
        if (op > 0):
            op = op + 1
            strRet = oURL[op:]
            strRet = strRet.lower()
            return strRet

if __name__ == "__main__":
    oHomePage = 'http://www.1ppt.com/kejian/'
    oSpider = CPPT1Spider(oHomePage)
    oSpider.GetIndexPage()