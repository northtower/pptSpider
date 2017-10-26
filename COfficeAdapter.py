#!/usr/bin/python
# -*- coding: gbk -*-

'''''
对MSOffice COM接口的封装，简单实现文件打开、操作和关闭等行为。
为数据清洗和快照提供支持。
'''

import win32com  
from win32com.client import Dispatch, constants  
import os

msoTrue = -1

class COfficeAdapter():
    def __init__(self ):
        print "init"
        self.m_App = win32com.client.Dispatch('PowerPoint.Application')
        #隐式加载为0 ，显式加载为1
        self.m_App.Visible = 1

    def __del__(self):
        print "__del__"
        self.m_App.Quit()

    def OpenDoc(self, oPath):
        oRet = False
        if os.path.exists(oPath):
            oPath = oPath.lower()
            oNameSplit = os.path.splitext(oPath)
            if (oNameSplit[1] == ".ppt" or oNameSplit[1] == ".pptx"):
                try:
                    self.m_Doc = self.m_App.Presentations.Open(oPath)
                    oRet = True
                except BaseException:
                    print "OpenDoc Exception!!"

        return oRet

    def SaveDoc(self, newfilename=None):
        if newfilename:       
            self.m_Doc.SaveAs(newfilename)                  
        else:   
            oRet = self.m_Doc.Save()
            print "save:" ,oRet

    def CloseDoc(self):
        #print self.m_Doc.Saved
        self.m_Doc.Close()

    # 删除文档中恶意广告信息，只要包含制定字符，即删除整个textRange。
    # 此方法，主要针对"www.1ppt.com"。
    # 有删除操作返回true
    def RemoveTextRange(self, oKeyStr):
        oRet = False
        oSlideCounts = self.m_Doc.Slides.Count
        #遍历每一页，方便对每页数据进行操作
        for i in range(1 , oSlideCounts + 1):
            #print i ,"页"
            oSlide = self.m_Doc.Slides.Item(i)
            oShapeCounts = oSlide.Shapes.Count
            #print oShapeCounts
            #遍历单页中所有shape
            for j in range(1,oShapeCounts + 1):
                oShape = oSlide.Shapes.Item(j)
                #判断类型，找到文字
                if oShape.TextFrame.HasText == msoTrue:
                    oTR = oShape.TextFrame.TextRange
                    try:
                        #oTextLen = oTR.Length
                        sText = oTR.Text
                        #转换为小写
                        sText = sText.lower()
                        #print sText

                        # 查找关键字，确认删除制定信息。
                        # 该方法并不用于最后一页，因为最后一页为整页的广告信息
                        oPos = sText.find(oKeyStr)
                        if oPos > 0  and i != oSlideCounts :
                            print "remove textRange" , i ,"+" ,oSlideCounts
                            oRet = True
                            oShape.Delete()
                            break
                        
                        if oPos > 0  and i == oSlideCounts :
                            print "delete Slide"
                            oRet = True
                            oSlide.Delete()
                            break                       

                    except BaseException:
                        print "BaseException"
        return  oRet


def testDoc():
    #oFilePath = r"D:\2.ppt"
    oFilePath = r"D:\3.pptx"
    #oFilePath = r"D:\pptSpider\PPTFile1\A版小学一年级上册语文PPT课件\2015语文A版语文一上《蜘蛛织网》ppt课件1.pptx"
    oKeyWork = "www.1ppt.com"
    print oFilePath
    oDoc = COfficeAdapter()
    if oDoc.OpenDoc(oFilePath):
        oRem = oDoc.RemoveTextRange(oKeyWork)
        #只有当删除页面元素时，才进行文档保存操作
        if oRem:
            oDoc.SaveDoc()
        oDoc.CloseDoc()

#testDoc()