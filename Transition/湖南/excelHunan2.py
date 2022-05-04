# -*- coding: utf-8 -*-
"""
Created on Sun Apr  3 12:12:12 2022
@author: TongJi|NWPU wjm
No pains,no gains. 
You must to have the ability to protect family mumbers.
"""

import pandas as pd 
import numpy as np
import re
import os
import xlwt
import xlrd

#参考数据 湖南省注册会计师2018年度任职资格检查合格人员名单（第一批）.xls
#----------------------------------
'''
目录
    数据预处理 ：在数据的最后一行后，新增一行，在第三个单元格填写"尾部"字段，该行会被识别为空行，便于程序提取事务所区块
    1.1 exceltoDataFrame 将excel中各事务所导入为一个DataFrame
        输出：保留excel视觉结构的一个DataFrame格式文件和记录下每个空行出现时在DataFrame中的索引号，以便1.2进行操作
    1.2 dataFrametoList 将DataFrame转化为是List格式 HunanExcel特征，除含有中文数字的城市信息的下一行也为空行外，其余事务所块之间必存在一个空行用于分割，为保证正常分割，
            需保证Excel以空行开始，以空行结束
        1.2.1 dropErrorData 去除城市行下的空行，将城市归属进某一事务所区块中（在后面会对含有城市和不含有城市的区块进行分类处理），保证 区块间有一个空行，区块内有一个空行，首行尾行均为空行
        输出 ：删除完空行的DataFrame
    1.3 dataFrametoList 将不同事务所存放在不同区块中，以list保存，此时区块中应仅有且必须仅有一行空行 （可自行计算区块内空行数量，如有异常，必须调整）
        输出 ：分割后的每一个事务所信息，以List存储
    1.4 dataTwiceIndex  第二次检测空行在每个事务所区块中的位置，用于分割事务所区块中非会计师姓名信息和会计师姓名信息的部分
        输出 : 分割
    1.5 columnData 提取非姓名信息
    1.6 findName  提取List中姓名列、ID列，分别保存为含有对应事务所名称列的DataFrame格式，合并对应的姓名与ID，最后根据事务所名称合并 事务所非会计师姓名信息与会计师姓名信息
    1.7 creatExcel 向目标保存路径保存为符合excel输出结构的DataFrame格式
        1.7.1 astype 将id列的整型转化为字符串格式
    1.8 removeNaN 去除事务所会计师姓名列中的空行
'''
#----------------------------------
def dropErrorData(dataframe): #去除省份下的空行
    liststr = ['一、','二、','三、','四、','五、','六、','七、','八、','九、','十、']  #识别城市行的标志，城市行必会带有中文数字，而第一列姓名中很难包括，为保证正确率，可将函数调整为根据对空行的下一行进行检测，则不会意外删除第一列姓名中含有数字的人所在行
    listindex = []
    for i in liststr:
        for j in range(dataframe.shape[0]):
            if bool(re.search(i,str(dataframe.iloc[j,0]))):
                listindex.append(j+1)
            else:
                continue
    print(listindex)
    dataframe = dataframe.drop(listindex).reset_index() # 执行删除操作后，对索引重新排序(DataFrame没有自动排序的功能，执行操作后依旧会保持原来的索引号，非连续的索引号不利于合并操作，因此进行重排序，并保证原有数据相对位置不变)
    return dataframe
    
def exceltoDataFrame(filename : str): #将excel中各事务所导入为一个DataFrame
    data = pd.read_excel(filename,sheet_name='年检合格人员名单 (第一批)  ',header=None,skiprows=[0,1,2,3,4]) #skiprows：跳过list中对应值的索引号所代表的行
    
    data = dropErrorData(data).drop('index',axis=1)
    listIndex = []
    for i in range(data.shape[0]):
        if(data.iloc[i,0] is np.nan and data.iloc[i,1] is np.nan):  #检测空行机制
            listIndex.append(i)
        else:
            continue
    
    return data,listIndex
    


def dataFrametoList(framename,numberOfIndex : list): #按照NaN所在位置索引对事务所进行分块，一个事务所为一个List元素
    splitData = []
    for i in range(len(numberOfIndex)):
        if ((i % 2 == 0) & (i != len(numberOfIndex)-1)):
            splitData.append(framename.iloc[numberOfIndex[i]+1:numberOfIndex[i+2],0:6]) #事务所的数据只存于前六列中
        else:
            continue
    return splitData

def dataTwiceIndex(splitOnce): #第二次检测空行在每个事务所区块中的位置，用于分割事务所区块中非会计师姓名信息和会计师姓名信息的部分
    index = []
    for i in range(len(splitOnce)):
        for j in range(splitOnce[i].shape[0]):
           if splitOnce[i].iloc[j,0] is np.nan:
                index.append(j)
           else:
                continue
    return index
    
def columnData(splitOnce,index):
    sname = [] #存放所有事务所名称
    substation = []#存放所有事务所的分类 总所or分所
    superintendent = []#存放所有事务所所长的姓名(如有)
    str1 = '分所' #可根据数据情况，扩充形容词
    str2 = '总所' #可根据数据情况，扩充形容词
    
    for i in range(len(splitOnce)): # 根据事务所区块的不同情况进行分类处理，非姓名信息在不同的区块中的相对位置不同
        if index[i] == 4 :
            sname.append(splitOnce[i].iloc[0,0])
            if(str1 in splitOnce[i].iloc[0,0]):
                substation.append(str1)
            else:
                substation.append(str2)
            if(splitOnce[i].iloc[1,3] is not np.nan):
                superintendent.append(splitOnce[i].iloc[1,3])
            else:
                superintendent.append("")
        elif index[i] == 5 :
            sname.append(splitOnce[i].iloc[1,0])
            if (str1 in splitOnce[i].iloc[1,0]):
                substation.append(str1)
            else:
                substation.append(str2)
            if(splitOnce[i].iloc[2,3] is not np.nan):
                superintendent.append(splitOnce[i].iloc[2,3])
            else:
                superintendent.append("")
        elif index[i] == 1:
            sname.append(splitOnce[i].iloc[0,0])
            substation.append("Null")
            superintendent.append("Null")
        else:
            sname.append("有问题")
    return sname,substation,superintendent

def findName(splitOnce,index):
    enameSet = pd.DataFrame({})
    eIdSet = pd.DataFrame({})
    for i in range(len(splitOnce)):
            enamelist = []
            eIdlist = []
            
            enamelist.append(splitOnce[i].iloc[index[i]+3:splitOnce[i].shape[0],0])
            enamelist.append(splitOnce[i].iloc[index[i]+3:splitOnce[i].shape[0],2])
            enamelist.append(splitOnce[i].iloc[index[i]+3:splitOnce[i].shape[0],4])
            enameDF = pd.concat([enamelist[0],enamelist[1],enamelist[2]],ignore_index=True)

            eIdlist.append(splitOnce[i].iloc[index[i]+3:splitOnce[i].shape[0],1])
            eIdlist.append(splitOnce[i].iloc[index[i]+3:splitOnce[i].shape[0],3])
            eIdlist.append(splitOnce[i].iloc[index[i]+3:splitOnce[i].shape[0],5])
            eIdDF = pd.concat([eIdlist[0],eIdlist[1],eIdlist[2]],ignore_index=True)
            
            sname = pd.DataFrame({splitOnce[i].iloc[0 if index[i]==4 or index[i]==1 else 1,0]})
            snameDF = pd.DataFrame({})
            for i in range(enameDF.shape[0]):
                snameDF = pd.concat([snameDF,sname]) 
            
            enameDF = pd.concat([snameDF.reset_index(),enameDF],axis=1,ignore_index=True)
            eIdDF = pd.concat([snameDF.reset_index(),eIdDF],axis=1,ignore_index=True)
            
            enameSet = pd.concat([enameSet,enameDF],axis=0)
            eIdSet = pd.concat([eIdSet,eIdDF],axis=0)
            
    enameSet = enameSet.drop(0,axis=1).reset_index()
    eIdSet = eIdSet.drop(0,axis=1).reset_index()
    enameSet = enameSet.drop('index',axis=1)
    eIdSet = eIdSet.drop('index',axis=1)
    enameSet.columns = ['事务所全称','注册会计师姓名']
    eIdSet.columns = ['事务所','注册会计师ID']
    
    enameSet = enameSet.join(eIdSet).drop('事务所',axis=1)
    
    return enameSet
            

def creatExcel(sname,substation,superintendent,enameSet):
    
    sname = pd.DataFrame(sname)
    sname.columns = ['事务所全称']
    substation = pd.DataFrame(substation)
    substation.columns = ['总所or分所']
    superintendent = pd.DataFrame(superintendent)
    superintendent.columns = ['所长/分所所长（如有）']
    excel = pd.concat([sname,substation,superintendent],axis=1)
    excel = pd.merge(excel, enameSet,left_on='事务所全称',right_on='事务所全称',sort=False).reset_index().drop('index',axis=1) #每次合并后 原有的索引值不会重新排序，需手动排序，原有索引会以'index'列名称存储，需要删除
    excel = astype(excel, '注册会计师ID')
    
    return excel

def removeNaN(dataframe,objectcolumn : str):
    print(dataframe[objectcolumn].isnull().value_counts())
    dataframe[objectcolumn] = dataframe[objectcolumn].fillna('空值')
    indexlist = dataframe[(dataframe[objectcolumn] =='空值')].index.tolist()
    dataframe = dataframe.drop(indexlist)
    return dataframe

def astype(dataframe,objectcolumn : str):
    dataframe = dataframe.astype({objectcolumn:'str'})
    return dataframe

def date(stamp):
    #输出时间格式
    
    delta = pd.Timedelta(str(stamp)+'D')
    real_time = pd.to_datetime('1899-12-30') + delta
    return real_time

def word_extraction_Hunan_table(province:str,year:str):
    
    path = r'E:\Spyderworkplace\excel\湖南' #输入文件所在文件夹
    path_output = r'E:\Spyderworkplace\excel\湖南' #输出文件所在文件夹
    filelist = os.listdir(path) #读取文件夹下所有文件，以list格式存储
    
    #生成excel的表头
    
    i_out = 1 #记录加载行数
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
    worksheet.write(0, 0, label='年度')
    worksheet.write(0, 1, label='通知发布日期')
    worksheet.write(0, 2, label='来源公告名称')
    worksheet.write(0, 3, label='事务所所在省份')
    worksheet.write(0, 4, label='事务所全称')
    worksheet.write(0, 5, label='总所or分所')
    worksheet.write(0, 6, label='总/分所所长（如有）')
    worksheet.write(0, 7, label='注册会计师姓名')
    worksheet.write(0, 8, label='注册会计师ID')
    worksheet.write(0, 9, label='备注')
    
    for file in filelist:
        
        if ('.xls' in file or '.xlsx' in file) and year in file and province in file and '第一批' in file:  # 搜索xlsx或xlsx文件
        
            word_path = os.path.join(path, file) #链接文件名和文件夹路径，形成完整路径
            
            data,index = exceltoDataFrame(word_path)
            splitlist = dataFrametoList(data, index)
            index2 = dataTwiceIndex(splitlist)
            sname1,substation1,superintendent1 = columnData(splitlist, index2)
            enameSet1= findName(splitlist, index2)
            excel1= creatExcel(sname1, substation1, superintendent1, enameSet1)
            excel1_final = removeNaN(excel1, '注册会计师姓名')
            
            wb_read = xlrd.open_workbook(os.path.join(path,'源地址.xls'))  # 从源地址excel提取'通知发布日期' '来源公告名称',要求该文件也在path文件夹下
            
            for i in range(wb_read.sheet_by_index(0).nrows - 1):
                if year in wb_read.sheet_by_index(0).cell_value(i, 0) and province in wb_read.sheet_by_index(
                        0).cell_value(i, 0):
                    d = wb_read.sheet_by_index(0).cell(i, 1).value
                    date_pb = str(date(d)).split(" ")[0]
                    name_pb = wb_read.sheet_by_index(0).cell_value(i, 0)


            for st in range(excel1_final.shape[0]):
                    
                    worksheet.write(i_out, 0, year)
                    worksheet.write(i_out, 1, date_pb)
                    worksheet.write(i_out, 2, name_pb)
                    worksheet.write(i_out, 3, province)
                    
                    #清洗事务所名称中无关字符
                    name = excel1_final.iloc[st,0]
                    name = re.sub('[a-zA-Z0-9’!"#$%&\'()*+,-./；:;<=>?@，：。?★、…【】《》？“”‘’！[\\]^_`{|}~\s一二三四五六七八九十]+', "",name)  # 事务所名称
                    name = re.sub('人', '',name)
                    name = re.sub('（）', '',name)
                    
                    worksheet.write(i_out, 4, name)  # 事务所名称
                    worksheet.write(i_out, 5, excel1_final.iloc[st,1]) #是否分所
                    worksheet.write(i_out,6, excel1_final.iloc[st,2]) #所长姓名
                    worksheet.write(i_out, 7, excel1_final.iloc[st,3]) #会计师姓名
                    worksheet.write(i_out,8,excel1_final.iloc[st,4]) #会计师ID
                    i_out = i_out + 1

                    
    workbook.save(os.path.join(path_output,province+year+'.xls'))

if __name__ == '__main__':
    
    
    #逐步测试函数
    file = '湖南省注册会计师2018年度任职资格检查合格人员名单（第一批）1.xls' 
    #导入路径 保证程序与导入路径在同一文件夹下，如在不同文件夹，需提供完整路径
    
    data,index = exceltoDataFrame(file)

    
    splitlist = dataFrametoList(data, index)
    
    index2 = dataTwiceIndex(splitlist)
    sname1,substation1,superintendent1 = columnData(splitlist, index2)
    enameSet1= findName(splitlist, index2)
    excel1= creatExcel(sname1, substation1, superintendent1, enameSet1)
    excel1_final = removeNaN(excel1, '注册会计师姓名')
    excel1_final.to_excel("数据筛选_{}".format(file)) #导出数据
    
    
    #测试完成，输出最终结果
    #word_extraction_Hunan_table('湖南省', str(2018)) #如果文件夹中存在 同省份同年份的不同批次文件，需修改 214行的文件判断条件
    