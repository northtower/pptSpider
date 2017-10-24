#!/usr/bin/python
# -*- coding: gbk -*-

import requests
import re
import os
from lxml import etree
from requests.exceptions import ReadTimeout,ConnectionError,RequestException
import rarfile

import sys
reload(sys)
sys.setdefaultencoding('gbk')
#sys.setdefaultencoding('utf-8')

import pyUnzip

def getFileName(oURL):
    op = oURL.rfind('/')
    if (op > 0):
        op = op + 1
        strRet = oURL[op:]
        strRet = strRet.lower()
        return strRet

def downLoadByURL(oUrl , dirPath):

    html = requests.get(oUrl).text
    selector = etree.HTML(html)
    zipUrl  = selector.xpath('//ul[@class="downurllist"]/li/a/@href')
    if not zipUrl:
        return  False

    strZip = str(zipUrl[0])
    print strZip

    oHeader = { "Accept":"text/html,application/xhtml+xml,application/xml;",
                "Accept-Encoding":"gzip",
                "Accept-Language":"zh-CN,zh;q=0.8",
                "Referer":"http://www.1ppt.com/",
                "User-Agent":"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.90 Safari/537.36"
                }

    PDFName = getFileName(strZip)
    localPDF = dirPath + "\\" + PDFName

    try:
        response = requests.get(strZip,headers = oHeader ,timeout=10)
        print(response.status_code)
        print localPDF

        #文件下载
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
    #print type(PDFName)
    #pyUnzip.unzipFile(PDFName)


def getContentByList(oListPageURL):

    html = requests.get(oListPageURL).text
    selector = etree.HTML(html)

    # listPage 的获取办法
    content3 = selector.xpath('//ul[@class="arclist"]/li/a/@href')
    for value in content3:
        print value

    # 找到首页和尾页，得到所有内容.
   # content4 = selector.xpath(u'//ul[@class="pages"]/li/a')

    #have nextpage 通过“下一页”的关键字查找
    #print u"//a[contains(text(), '下一页')]"
    content4 = selector.xpath(u"//a[contains(text(), '下一页')]")
    for value in content4:
        print value.text

    #decode("utf-8").encode('gbk')

#ppt1_url = 'http://www.1ppt.com/kejian/29032.html'
#downLoadByURL(ppt1_url)


#oUrl = 'http://www.1ppt.com/kejian/yuwen/138/kejian_1.html'
#getContentByList(oUrl)