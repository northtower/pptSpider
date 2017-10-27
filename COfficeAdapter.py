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
        try:
            self.m_App = win32com.client.Dispatch('PowerPoint.Application')
            self.m_App.Visible = 1
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


def testDoc():
    oFilePath = r"D:\pptSpider\PPTFile1\A��Сѧ���꼶�²�����PPT�μ�\2015����A���������¡�һ˫��Ь��ppt�μ�.pptx"
    oKeyWork = "www.1ppt.com"
    print oFilePath
    oDoc = COfficeAdapter()
    if oDoc.OpenDoc(oFilePath):
        oRem = oDoc.RemoveTextRange(oKeyWork)
        #ֻ�е�ɾ��ҳ��Ԫ��ʱ���Ž����ĵ��������
        if oRem:
            oDoc.SaveDoc()
        oDoc.CloseDoc()

#testDoc()