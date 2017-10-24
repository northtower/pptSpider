# -*- coding:utf8 -*-

import requests
from lxml import etree

import sys
reload(sys)
sys.setdefaultencoding('gbk')
#sys.setdefaultencoding('utf-8')

import PS_downloadFile

#oListPage = "http://www.1ppt.com/kejian/yuwen/161/"

def getContentByList(oListPageURL , oLocalPath):

    homePage = "http://www.1ppt.com"
    html = requests.get(oListPageURL)
    html.encoding = 'gb2312'
    selector = etree.HTML(html.text)

    # listPage 的获取办法
    content3 = selector.xpath('//ul[@class="arclist"]/li/a/@href')
    for value in content3:
        listPageUrl = homePage + value
        print listPageUrl
        oRet = PS_downloadFile.downLoadByURL(listPageUrl , oLocalPath)
        if not oRet:
            print "try again!!"
            PS_downloadFile.downLoadByURL(listPageUrl , oLocalPath)


    #have nextpage 通过“下一页”的关键字查找
    #print u"//a[contains(text(), '下一页')]"
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
        getContentByList(nextPage , oLocalPath)


#oStartPage = "kejian_1.html"
#oUrl = oListPage + oStartPage
#getContentByList(oUrl)