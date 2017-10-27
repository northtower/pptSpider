#!/usr/bin/python
# -*- coding: gbk -*-


import os
from os import path
import COfficeAdapter

RootDir   = "D:\\pptSpider\\PPTFile"
oKeyWork = "www.1ppt.com"

def doSomething(oRootDir):
    fileCounts = 1
    dirCounts  = 0
    for rt, dirs, files in os.walk(oRootDir):
        try:
            strPath = unicode(rt, "GB2312")
        except BaseException:
            print "strPath Exception!!"
            pass
        op = strPath.rfind('\\')
        dirCounts = dirCounts + 1
        if (op > 0):
            op = op + 1
            # 拿到目录名
            dirName = strPath[op:]
            print "[" ,dirCounts,"]" ,dirName

        #if dirCounts < 265 :
        #    continue

        oMSA = COfficeAdapter.COfficeAdapter()
        for fileName in files:

            '''
            createPath = OutputDir + dirName
            if not os.path.exists(createPath):
                os.mkdir(createPath)
            '''

            #拿到文件名和类型
            try:
                fileName = unicode(fileName, "GB2312")
                fname = os.path.splitext(fileName)

                #拼接输出文件名
                outZipFile = strPath + "\\" + fname[0] + ".zip"

                #getFileType
                oFileType = fname[1]
                oFileType = oFileType.lower()
                #fileCounts = fileCounts + 1
                #if (oFileType == ".ppt" or oFileType == ".pptx" ) and not path.isfile(outZipFile):
                if oFileType == ".ppt" or oFileType == ".pptx":
                    fileCounts = fileCounts + 1
                    #文档找到，进行格式转换
                    openFilePath = strPath + "\\" + fileName
                    print openFilePath

                    if oMSA.OpenDoc(openFilePath):
                        oRemObj = oMSA.RemoveTextRange(oKeyWork)
                        if oRemObj:
                            oMSA.SaveDoc()
                        oMSA.CloseDoc()

                    #print outZipFile
            except BaseException:
                print "splitext Exception!!"
                oMSA.CloseDoc()
                pass
        del oMSA

    print fileCounts

doSomething(RootDir)