# -*- coding:utf8 -*-

import pandas as pd
import numpy as np

'''''
Pandas相关的一些方法，提供输入的导入导出。
为之后的数据分析做准别
'''

class CPandasUtility():
    def __init__(self ):
        print "CPandasUtility __init__"
    
    def __del__(self):
        print "CPandasUtility __del__"

        
    def csvInfo(self):        
        filePath = r"D:\pptSpider\ex1.csv"
        jsonPath = r"D:\pptSpider\ex1.json"
        df = pd.read_csv(filePath)
        #print df

        print "shape" , df.shape
        print "Length" , len(df)

        print "head" , df.head
        print ""

        print "describe" , df.describe()

        #计算总和
        print "sum" , df.sum()

        #Series
        co1 =df["pageCounts"]
        print "type df" , type(df)
        print "type col" , type(co1)
        print "col:" 
        print co1

        print "Series shape:" , co1.shape
        print "Series index:" , co1.index
        print "Series values:" , co1.values
        print "Series name:" , co1.name
        #计算平均数
        #print "mean" , df.mean()
        '''
        print "Column Headers" , df.columns
        #print "data types:" , df.dtypes

        #print "info" , df.info
        print "Count" , df.count

        print "index" , df.index
        #print "Values" . df.values

        print df.sort_values("pageCounts").head(3)

        df.to_json(jsonPath)
        '''


    
def craeteDF():
    
    #pd.set_option('display.width',200)

    dates = pd.date_range('20130101', periods=6)
    df = pd.DataFrame(np.random.randn(6,4), index=dates, columns=list('ABCD'))

    docInfo = {
        'FileName' : "大家看得见的.ppt",
        'FontName' : "黑体",
        'SlideItem' : 1,
        'Index' : 1,
        'Text' : "打开端口打开的",
        'FontSize' : 23 }
    df3 = pd.DataFrame(docInfo , index=[1])
    
    print df3

    s2 = pd.Series({
        'FileName' : "newDoc.pptt",
        'FontName' : "黑体",
        'SlideItem' : 1,
        'Index' : 1,
        'Text' : "ssssss",
        'FontSize' : 21 })
    #s2=pd.Series(np.array(["newDoc.ppt","黑体" ,6,7,"打开端口打开的" ,8]))
    print s2

    df3 = df3.append(s2 , ignore_index=True)
    print df3

    #print df3

    #s1=pd.Series(np.array(["FileName","SlideItem","Index","Text" , "FontName", "FontSize"]))
    
    #df4=pd.DataFrame([s1,s2])
    #print df4


    #print "Column Headers" , df2.columns
    #print "shape num:" ,df2.shape[0] + 1
    #df2.loc[df2.shape[0]] = {"123.ppt" , 5 , 1 , "text" , "NewFontName" , 10}
    #print df2

    #df3 = df2.append(newObj , ignore_index=True)
    #print "df3"
    #print df3
    
    #查看不同列的数据类型
    #print df2.dtypes

    #通过标签来在多个轴上进行选择
    #print df2.loc[:,["FileName" , "Index"]]

    #描述性统计
    #print df2.mean()

    #append

    #filePath = r"D:\pptSpider\ex2.csv"
    #df2.to_csv(filePath)



craeteDF()