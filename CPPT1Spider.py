# -*- coding:utf8 -*-

import requests
import os
import xlsxwriter

from lxml import etree

LOCALPATH = "/Users/Sophia/Codes/python/spider/pptDownLoad/"
HOMEPAGE  = 'http://www.1ppt.com'

class CPPT1Spider():
    def __init__(self , homePage):
        print "init"
        self.initWithHomePage(homePage)
    
    def initWithHomePage(self , oHomePage):
        print ""
        html = requests.get(oHomePage)
        html.encoding = 'gb2312'
        selector = etree.HTML(html.text)

        self.m_PageListDes = selector.xpath('//dd[@class="ikejian_col_nav"]/ul/li/h4/a/@title')
        self.m_PageListURL = selector.xpath('//dd[@class="ikejian_col_nav"]/ul/li/h4/a/@href')


    def getContentByList(self , oListPageURL , oLocalPath):

        homePage = "http://www.1ppt.com"
        html = requests.get(oListPageURL)
        html.encoding = 'gb2312'
        selector = etree.HTML(html.text)

        # listPage 的获取办法
        content3 = selector.xpath('//ul[@class="arclist"]/li/a/@href')
        for value in content3:
            listPageUrl = homePage + value
            print listPageUrl
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
            self.getContentByList(nextPage , oLocalPath)


    def GetContentList(self):   
        oCounts = 1             
        for value1,value2 in zip(self.m_PageListDes ,self.m_PageListURL):
            oListPageURL = HOMEPAGE + value2
            createPath = LOCALPATH + value1
            print "[", oCounts, "]", value1, oListPageURL
            oCounts = oCounts + 1        
            self.getContentByList(oListPageURL , createPath) 
            #if not os.path.exists(createPath):
                #os.mkdir(createPath)
                #PS_getContentPage.getContentByList(oListPageURL , createPath)
            #print createPath

    def getFileList(self):
        raise ""


if __name__ == "__main__":
    print "main"
    oHomePage = 'http://www.1ppt.com/kejian/'
    oSpider = CPPT1Spider(oHomePage)
    oSpider.GetContentList()

#oHomepage = 'http://www.1ppt.com/kejian/'
#getHomePage(oHomepage)