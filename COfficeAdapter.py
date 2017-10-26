#!/usr/bin/python
# -*- coding: utf-8 -*-
'''''
    对MSOffice COM接口的封装，简单实现文件打开、操作和关闭等行为。
    为数据清洗和快照提供支持。
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


oFilePath = r"D:\2.ppt"
#oFilePath = r"D:\pptSpider\PPTFile\北师大版小学一年级上册数学PPT课件\1.3《小猫钓鱼》课件.ppt"
oDoc = COfficeAdapter()
oDoc.OpenDoc(oFilePath)
oDoc.CloseDoc()
