#!/usr/bin/python
# -*- coding: gbk -*-

import requests
import re
from lxml import etree
import rarfile

import sys
reload(sys)
sys.setdefaultencoding('gbk')
#sys.setdefaultencoding('utf-8')


def getFileName(oURL):
    op = oURL.rfind('/')
    if (op > 0):
        op = op + 1
        strRet = oURL[op:]
        strRet = strRet.lower()
        return strRet

def downLoadByURL(oUrl):

    html = requests.get(oUrl).text
    selector = etree.HTML(html)
    zipUrl  = selector.xpath('//ul[@class="downurllist"]/li/a/@href')
    strZip = str(zipUrl[0])
    print strZip

    oHeader = { "Accept":"text/html,application/xhtml+xml,application/xml;",
                "Accept-Encoding":"gzip",
                "Accept-Language":"zh-CN,zh;q=0.8",
                "Referer":"http://www.1ppt.com/",
                "User-Agent":"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.90 Safari/537.36"
                }

    req = requests.get(strZip,headers = oHeader)
    print req

    localDir = 'D:\\pptSpider\\downloadFile\\'
    PDFName = getFileName(strZip)
    localPDF = localDir + PDFName
    print localPDF

    oFile = open(localPDF, "wb")
    for chunk in req.iter_content(chunk_size=512):
        if chunk:
            oFile.write(chunk)

    unzipFile(PDFName)


def unzipFile(fileName):
    oFilePath = "D:\\pptSpider\\downloadFile\\" + fileName
    fantasy_zip = rarfile.RarFile(oFilePath)
    fantasy_zip.extractall("D:\\pptSpider\\downloadFile")

    '''
    for oFileName in fantasy_zip.namelist():
        filenameLower = oFileName.lower()
        if filenameLower.find('.ppt') > 1:
            print "find it:[" + oFileName + "]"
            #fantasy_zip.extract(oFileName , "D:\\pptSpider\\downloadFile")
            fantasy_zip.extractall(u"D:\\pptSpider\\downloadFile")
    '''
    fantasy_zip.close()

def getContentByList(oListPageURL):

    html = requests.get(oListPageURL).text
    selector = etree.HTML(html)

    # listPage 的获取办法
    content3 = selector.xpath('//ul[@class="arclist"]/li/a/@href')
    #for value in content3:
    #    print value

    # 找到首页和尾页，得到所有内容.
    content4 = selector.xpath(u'//ul[@class="pages"]/li/a')

    #have nextpage 通过“下一页”的关键字查找
    #print u"//a[contains(text(), '下一页')]"
    #content4 = selector.xpath(u"//a[contains(text(), '下一页')]")
    for value in content4:
        print value.text

    #decode("utf-8").encode('gbk')

ppt1_url = 'http://www.1ppt.com/kejian/29032.html'
downLoadByURL(ppt1_url)
#unzipFile()

#oUrl = 'http://www.1ppt.com/kejian/yuwen/138/kejian_1.html'
#getContentByList(oUrl)