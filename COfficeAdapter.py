#!/usr/bin/python
# -*- coding: gbk -*-

'''''
��MSOffice COM�ӿڵķ�װ����ʵ���ļ��򿪡������͹رյ���Ϊ��
Ϊ������ϴ�Ϳ����ṩ֧�֡�
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
        #��ʽ����Ϊ0 ����ʽ����Ϊ1

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


    # ɾ���ĵ��ж�������Ϣ��ֻҪ�����ƶ��ַ�����ɾ������textRange��
    # �˷�������Ҫ���"www.1ppt.com"��
    # ��ɾ����������true
    def RemoveTextRange(self, oKeyStr):
        oRet = False
        oSlideCounts = self.m_Doc.Slides.Count
        #����ÿһҳ�������ÿҳ���ݽ��в���
        for i in range(1 , oSlideCounts + 1):
            #print i ,"ҳ"
            oSlide = self.m_Doc.Slides.Item(i)
            oShapeCounts = oSlide.Shapes.Count
            #print oShapeCounts
            #������ҳ������shape
            for j in range(1,oShapeCounts + 1):
                oShape = oSlide.Shapes.Item(j)
                #�ж����ͣ��ҵ�����
                if oShape.TextFrame.HasText == msoTrue:
                    oTR = oShape.TextFrame.TextRange
                    try:
                        #oTextLen = oTR.Length
                        sText = oTR.Text
                        #ת��ΪСд
                        sText = sText.lower()
                        #print sText

                        # ���ҹؼ��֣�ȷ��ɾ���ƶ���Ϣ��
                        # �÷��������������һҳ����Ϊ���һҳΪ��ҳ�Ĺ����Ϣ
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

    # ��ȡslide��shape������
    def GetShapeInfo(self):
        oRet = False
        oSlideCounts = self.m_Doc.Slides.Count

        #����ÿһҳ�������ÿҳ���ݽ��в���
        for i in range(1 , oSlideCounts + 1):
            print "��", i ,"ҳ"
            oSlide = self.m_Doc.Slides.Item(i)
            oShapeCounts = oSlide.Shapes.Count
            #print oShapeCounts
            #������ҳ������Ԫ�أ����ͷ�Χ�ǣ�Chart��SmartArt��Table��TextFrame��
            for j in range(1,oShapeCounts + 1):
                oShape = oSlide.Shapes.Item(j)
                #�ж����ͣ��ҵ�����
                print "Name:" , oShape.Name                
                print "Type:" , oShape.Type
                #if oShape.Type == 1 :
                #    print "Name:" , oShape.Name
                                    
                try:
                    if oShape.TextFrame.HasText == msoTrue :
                        oTR = oShape.TextFrame.TextRange              

                        #���塢�ֺ�      
                        oFont = oShape.TextFrame.TextRange.Font
                        #�ı�����
                        oOrientation = oShape.TextFrame.Orientation 

                        #ˮƽ���뷽ʽ
                        oHAnchor = oShape.TextFrame.HorizontalAnchor  
                        oVAnchor = oShape.TextFrame.VerticalAnchor   

                        #�ı����λ�úʹ�С���԰�Ϊ��λ
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

                        ##������ı���Ϣ���ж��������͡�����ͨ�ı����������
                        '''
                        if oShape.Type == 1 :
                            oPos = oShape.Name.find("WordArt")
                            print "oPos:",oPos
                            if oPos > 0 :
                                #�ж��Ƿ�Ϊ������
                                print "������"
                        '''

                    elif oShape.HasChart == msoTrue: #ͼ��
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
    oFilePath = r"D:\pptSpider\2015����A���������ϡ�һŵǧ��ppt�μ�.pptx"
    #oFilePath = r"D:\pptSpider\placeHolder.pptm"
    oKeyWork = "www.1ppt.com"
    print oFilePath
    oDoc = COfficeAdapter()
    if oDoc.OpenDoc(oFilePath):
        #oRem = oDoc.RemoveTextRange(oKeyWork)
        oRem = oDoc.GetShapeInfo()
        #ֻ�е�ɾ��ҳ��Ԫ��ʱ���Ž����ĵ��������
        if oRem:
            oDoc.SaveDoc()
        oDoc.CloseDoc()

testDoc()