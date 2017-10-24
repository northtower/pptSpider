# -*- coding:utf8 -*-

import requests
import os
from lxml import etree
import PS_getContentPage

homePage = "http://www.1ppt.com"
localPath = "D:\\pptSpider\\downloadFile\\"

#根据首页，获取各个学科的分页目录
def getHomePage(oListPageURL):

    html = requests.get(oListPageURL)
    html.encoding = 'gb2312'
    selector = etree.HTML(html.text)

    oCounts = 1
    content1 = selector.xpath('//dd[@class="ikejian_col_nav"]/ul/li/h4/a/@title')
    content2 = selector.xpath('//dd[@class="ikejian_col_nav"]/ul/li/h4/a/@href')
    for value1,value2 in zip(content1 ,content2):
        oListPageURL = homePage + value2
        oCounts = oCounts + 1
        createPath = localPath + value1
        if oCounts > 5:
            print "[", oCounts, "]", value1, oListPageURL
            if not os.path.exists(createPath):
                os.mkdir(createPath)
            PS_getContentPage.getContentByList(oListPageURL , createPath)

        #print createPath
        #

oHomepage = 'http://www.1ppt.com/kejian/'
getHomePage(oHomepage)