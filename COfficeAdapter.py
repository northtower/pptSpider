#!/usr/bin/python
# -*- coding: gbk -*-

'''''
��MSOffice COM�ӿڵķ�װ����ʵ���ļ��򿪡������͹رյ���Ϊ��
Ϊ������ϴ�Ϳ����ṩ֧�֡�
'''

import win32com  
from win32com.client import Dispatch, constants  
import os

msoTrue = -1

class COfficeAdapter():
    def __init__(self ):
        print "init"
        self.m_App = win32com.client.Dispatch('PowerPoint.Application')
        #��ʽ����Ϊ0 ����ʽ����Ϊ1
        self.m_App.Visible = 1

    def __del__(self):
        print "__del__"
        self.m_App.Quit()

    def OpenDoc(self, oPath):
        if os.path.exists(oPath):
            oPath = oPath.lower()
            oNameSplit = os.path.splitext(oPath)
            if (oNameSplit[1] == ".ppt" or oNameSplit[1] == ".pptx"):
                  self.m_Doc = self.m_App.Presentations.Open(oPath)

    def SaveDoc(self, newfilename=None):
        if newfilename:       
            self.m_Doc.SaveAs(newfilename)                  
        else:   
            self.m_Doc.Save() 

    def CloseDoc(self):
        print self.m_Doc.Saved
        self.m_Doc.Close()

    #ɾ���ĵ��ж�������Ϣ��ֻҪ�����ƶ��ַ�����ɾ������textRange��
    #�˷�������Ҫ���"www.1ppt.com"��
    def RemoveTextRange(self, oKeyStr):
        oSlideCounts = self.m_Doc.Slides.Count
        #����ÿһҳ�������ÿҳ���ݽ��в���
        for i in range(1 , oSlideCounts + 1):
            print i ,"ҳ"
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
                        print sText

                        # ���ҹؼ��֣�ȷ��ɾ���ƶ���Ϣ��
                        # �÷��������������һҳ����Ϊ���һҳΪ��ҳ�Ĺ����Ϣ
                        oPos = sText.find(oKeyStr)
                        if oPos > 0  and i != oSlideCounts :
                            print "remove textRange"
                            oShape.Delete()                      
                        
                        if oPos > 0  and i == oSlideCounts :
                            print "delete Slide"
                            oSlide.Delete()
                            break                       

                    except BaseException:
                        print "BaseException"


def testDoc():
    oFilePath = r"D:\pptSpider\PPTFile\��ʦ���Сѧһ�꼶�ϲ���ѧPPT�μ�\1.3��Сè���㡷�μ�.ppt"
    oKeyWork = "www.1ppt.com"
    print oFilePath
    oDoc = COfficeAdapter()
    oDoc.OpenDoc(oFilePath)
    oDoc.RemoveTextRange(oKeyWork)
    oDoc.SaveDoc()
    oDoc.CloseDoc()

testDoc()