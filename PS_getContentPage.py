# -*- coding:utf8 -*-

import requests
from lxml import etree

import sys
reload(sys)
#sys.setdefaultencoding('gbk')
sys.setdefaultencoding('utf-8')

import PS_downloadFile

'''
# 找到首页和尾页，得到所有内容.
    #content4 = selector.xpath(u'//ul[@class="pages"]/li/a')
'''
def getContentByList(oListPageURL):

    homePage = "http://www.1ppt.com"
    html = requests.get(oUrl)
    html.encoding = 'gb2312'
    selector = etree.HTML(html.text)

    # listPage 的获取办法
    content3 = selector.xpath('//ul[@class="arclist"]/li/a/@href')
    for value in content3:
        listPageUrl = homePage + value
        print listPageUrl
        PS_downloadFile.downLoadByURL(listPageUrl)


    #have nextpage 通过“下一页”的关键字查找
    #print u"//a[contains(text(), '下一页')]"
    content4 = selector.xpath(u"//a[contains(text(), '下一页')]")
    for value in content4:
        strlistUrl = value.get('href')
        print strlistUrl




oUrl = 'http://www.1ppt.com/kejian/yuwen/138/kejian_1.html'
getContentByList(oUrl)