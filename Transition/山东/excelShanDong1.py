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

def exceltoDataFrame(filename : str):
    data = pd.read_excel(filename,sheet_name='Sheet1',header=None)
    listIndex = []
    for i in range(data.shape[0]):
        if(data.iloc[i,0] is np.nan and data.iloc[i,1] is np.nan and data.iloc[i,2] is np.nan):
            listIndex.append(i)
        else:
            continue
    return data,listIndex


def dataFrametoList(framename,numberOfIndex : list):
    splitData = []
    for i in range(len(numberOfIndex)):
        if (i != len(numberOfIndex)-1) :
            splitData.append(framename.iloc[numberOfIndex[i]+1:numberOfIndex[i+1],0:8])
        else:
            continue
    return splitData

def dataTwiceIndex(splitOnce):
    index = []
    indexstr = '注册会计师'
    for i in range(len(splitOnce)):
        for j in range(splitOnce[i].shape[0]):
           if bool(re.search(indexstr,str(splitOnce[i].iloc[j,0]))):
                index.append(j)
           else:
                continue
    return index
    
def columnData(splitOnce,index):
    sname = []
    substation = []
    superintendent = []
    str1 = '分所'
    str2 = '总所'
    for i in range(len(splitOnce)):
        if index[i] == 4 :
            sname.append(splitOnce[i].iloc[1,2])
            if(str1 in splitOnce[i].iloc[1,2]):
                substation.append(str1)
            else:
                substation.append(str2)
            if(splitOnce[i].iloc[2,2] is not np.nan):
                superintendent.append(splitOnce[i].iloc[2,2])
            else:
                superintendent.append("Null")
        elif index[i] == 5 :
            sname.append(splitOnce[i].iloc[2,2])
            if (str1 in splitOnce[i].iloc[2,2]):
                substation.append(str1)
            else:
                substation.append(str2)
            if(splitOnce[i].iloc[3,2] is not np.nan):
                superintendent.append(splitOnce[i].iloc[3,2])
            else:
                superintendent.append("Null")
        elif index[i] == 1:
            sname.append(splitOnce[i].iloc[0,0])
            if(str1 in splitOnce[i].iloc[0,0]):
                substation.append(str1)
            else:
                substation.append(str2)
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
            
            enamelist.append(splitOnce[i].iloc[index[i]+2:splitOnce[i].shape[0],0])
            enamelist.append(splitOnce[i].iloc[index[i]+2:splitOnce[i].shape[0],2])
            enamelist.append(splitOnce[i].iloc[index[i]+2:splitOnce[i].shape[0],4])
            enamelist.append(splitOnce[i].iloc[index[i]+2:splitOnce[i].shape[0],6])
            enameDF = pd.concat([enamelist[0],enamelist[1],enamelist[2],enamelist[3]],ignore_index=True)

            eIdlist.append(splitOnce[i].iloc[index[i]+2:splitOnce[i].shape[0],1])
            eIdlist.append(splitOnce[i].iloc[index[i]+2:splitOnce[i].shape[0],3])
            eIdlist.append(splitOnce[i].iloc[index[i]+2:splitOnce[i].shape[0],5])
            eIdlist.append(splitOnce[i].iloc[index[i]+2:splitOnce[i].shape[0],7])
            eIdDF = pd.concat([eIdlist[0],eIdlist[1],eIdlist[2],eIdlist[3]],ignore_index=True)
            
            sname = pd.DataFrame({splitOnce[i].iloc[1 if index[i]== 4 else ( 5 if index[i]==5 else 0),2 if index[i]!=1 else 0]})
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
    excel = pd.merge(excel, enameSet,left_on='事务所全称',right_on='事务所全称',sort=False).reset_index().drop('index',axis=1)
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

if __name__ == '__main__':
    file = '2016年山东省注册会计师任职资格检查合格名单（第二批）.xlsx'
    data,index = exceltoDataFrame(file)
    splitlist = dataFrametoList(data, index)
    index2 = dataTwiceIndex(splitlist)
    sname1,substation1,superintendent1 = columnData(splitlist, index2)
    enameSet1= findName(splitlist, index2)
    excel1= creatExcel(sname1, substation1, superintendent1, enameSet1)
    excel1_final = removeNaN(excel1, '注册会计师姓名')
    excel1_final.to_excel("数据筛选_{}".format(file))