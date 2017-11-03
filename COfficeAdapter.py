#!/usr/bin/python
# -*- coding: gbk -*-

'''''
对MSOffice COM接口的封装，简单实现文件打开、操作和关闭等行为。
为数据清洗和快照提供支持。
'''

import win32com  
from win32com.client import Dispatch, constants  
import os
import sys
import CPandas

msoTrue = -1

#MsoShapeType
msoShapeTypeMixed = -2,
msoAutoShape = 1,
msoCallout = 2,
msoChart = 3,
msoComment = 4,
msoFreeform = 5,
msoGroup = 6,
msoEmbeddedOLEObject = 7,
msoFormControl = 8,
msoLine = 9,
msoLinkedOLEObject = 10,
msoLinkedPicture = 11,
msoOLEControlObject = 12,
msoPicture = 13,
msoPlaceholder = 14,
msoTextEffect = 15,
msoMedia = 16,
msoTextBox = 17,
msoScriptAnchor = 18,
msoTable = 19,
msoCanvas = 20,
msoDiagram = 21,
msoInk = 22,
msoInkComment = 23,
msoSmartArt = 24,
msoSlicer = 25,
msoWebVideo = 26,
msoContentApp = 27,
msoGraphic = 28,
msoLinkedGraphic = 29


class COfficeAdapter():
    def __init__(self ):
        print "init"
        try:
            self.m_App = win32com.client.Dispatch('PowerPoint.Application')
            self.m_App.Visible = 1
            self.m_Pandas = CPandas.CPandasUtility()
        except BaseException:
            print "init Exception!!"
        #隐式加载为0 ，显式加载为1

    def __del__(self):
        print "__del__"
        try:
            self.m_App.Quit()
        except BaseException:
            print "Quit Exception!!"


    def OpenDoc(self, oPath):
        oRet = False
        if os.path.exists(oPath):
            oPath = oPath.lower()
            oNameSplit = os.path.splitext(oPath)
            oFileName = os.path.split(oPath)
            if (oNameSplit[1] == ".ppt" or oNameSplit[1] == ".pptx"):
                try:
                    self.m_Doc = self.m_App.Presentations.Open(oPath)
                    self.m_Pandas.SetFileName(oFileName[1])
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
        try:
            self.m_Doc.Close()
        except BaseException:
            print "Close Exception!!"


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

    # 获取slide内shape的属性
    def GetShapeInfo(self):
        oRet = False
        oSlideCounts = self.m_Doc.Slides.Count

        #遍历每一页，方便对每页数据进行操作
        for i in range(1 , oSlideCounts + 1):
            print "第", i ,"页"
            oSlide = self.m_Doc.Slides.Item(i)
            oShapeCounts = oSlide.Shapes.Count
            #print oShapeCounts
            #遍历单页中所有元素，类型范围是（Chart、SmartArt、Table、TextFrame）
            for j in range(1,oShapeCounts + 1):
                oShape = oSlide.Shapes.Item(j)
                #判断类型，找到文字
                print "Name:" , oShape.Name                
                print "Type:" , oShape.Type
                #if oShape.Type == 1 :
                #    print "Name:" , oShape.Name
                                    
                try:
                    if oShape.TextFrame.HasText == msoTrue :
                        oTR = oShape.TextFrame.TextRange              

                        #字体、字号      
                        oFont = oShape.TextFrame.TextRange.Font
                        #文本方向
                        oOrientation = oShape.TextFrame.Orientation 

                        #水平对齐方式
                        oHAnchor = oShape.TextFrame.HorizontalAnchor  
                        oVAnchor = oShape.TextFrame.VerticalAnchor   

                        #文本框的位置和大小，以磅为单位
                        oBT = oShape.TextFrame.TextRange.BoundTop  
                        oBL = oShape.TextFrame.TextRange.BoundLeft
                        oBH = oShape.TextFrame.TextRange.BoundHeight 
                        oBW = oShape.TextFrame.TextRange.BoundWidth   

                        oTextLen = oTR.Length
                        sText = oTR.Text
                        print j , "-TextFrame:" ,sText 
                        print "[Len]:" ,oTextLen
                        #print "[Font.Name]:" ,oFont.Name ,"[Font.Size]:" ,oFont.Size
                        #print "Position:",oBT,":" ,oBL

                        self.m_Pandas.AppendText( i , j,oFont.Name,oFont.Size ,sText)

                        ##如果是文本信息，判断它的类型。如普通文本框和艺术字
                        '''
                        if oShape.Type == 1 :
                            oPos = oShape.Name.find("WordArt")
                            print "oPos:",oPos
                            if oPos > 0 :
                                #判断是否为艺术字
                                print "艺术字"
                        '''

                    elif oShape.HasChart == msoTrue: #图表
                        print j , "-Chart:"
                        
                    elif oShape.HasSmartArt  == msoTrue:
                        print j , "-SmartArt :"

                    elif oShape.HasTable   == msoTrue:
                        print j , "-Table :" 

                except BaseException:
                        exc_type, exc_obj, exc_tb = sys.exc_info()
                        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                        print " Get Elements Exception:" ,(exc_type, fname, exc_tb.tb_lineno)
        
        self.m_Pandas.PrintDataFrame()
        self.m_Pandas.SaveCsv()
        return  oRet

def testDoc():
    oFilePath = r"D:\pptSpider\2015语文A版语文六上《一诺千金》ppt课件.pptx"
    #oFilePath = r"D:\pptSpider\placeHolder.pptm"
    oKeyWork = "www.1ppt.com"
    print oFilePath
    oDoc = COfficeAdapter()
    if oDoc.OpenDoc(oFilePath):
        #oRem = oDoc.RemoveTextRange(oKeyWork)
        oRem = oDoc.GetShapeInfo()
        #只有当删除页面元素时，才进行文档保存操作
        if oRem:
            oDoc.SaveDoc()
        oDoc.CloseDoc()

testDoc()