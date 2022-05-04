# -*- coding: utf-8 -*-
"""
Created on Thu Apr  7 16:10:20 2022
@author: TongJi|NWPU wjm
No pains,no gains. 
You must to have the ability to protect family mumbers.
"""
import re

def splitNullRow(file1,file2):
    file1 = open(file1,'r',encoding = 'utf-8')
    file2 = open(file2,'w',encoding = 'utf-8')
    for line in file1.readlines():
        if line == '\n':
           line = line.strip("\n")
        if re.search(r'\.[0-9]+\.', line):
           #line = line.strip(re.search(r'\.[0-9]+\.', line).group(0))
           line = line.strip(re.search(r'[0-9]+', line).group(0)) #适用于上海2015
        file2.write(line)
    file1.close()
    file2.close()

if __name__ == '__main__':
    path1 = r'上海市2015年度通过任职资格检查注册会计师名单.txt'
    path2 = r'上海市2015年度通过任职资格检查注册会计师名单.txt_修改版.txt'
    splitNullRow(path1,path2)