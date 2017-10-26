#!/usr/bin/python
# -*- coding: gbk -*-

'''''
��MSOffice COM�ӿڵķ�װ����ʵ���ļ��򿪡������͹رյ���Ϊ��
Ϊ������ϴ�Ϳ����ṩ֧�֡�
'''

import win32com  
from win32com.client import Dispatch, constants  
import os

class COfficeAdapter():
    def __init__(self ):
        print "init"
        self.m_App = win32com.client.Dispatch('PowerPoint.Application')
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

    def CloseDoc(self):
        print self.m_Doc.Saved
        self.m_Doc.Close()

#oFilePath = r"D:\2.ppt"
oFilePath = r"D:\pptSpider\PPTFile\��ʦ���Сѧһ�꼶�ϲ���ѧPPT�μ�\1.3��Сè���㡷�μ�.ppt"
print oFilePath
oDoc = COfficeAdapter()
oDoc.OpenDoc(oFilePath)
oDoc.CloseDoc()
