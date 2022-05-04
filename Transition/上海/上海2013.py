import os
from docx import Document
import xlwt
import xlrd
import re
import pandas as pd

path_output = r'E:\Spyderworkplace\excel\修正数据'

def splitNullRow(file1):
    path2 = file1+str('修正')
    file1 = open(file1,'r',encoding = 'utf-8')
    file2 = open(path2,'w',encoding = 'utf-8')
    for line in file1.readlines():
        if line == '\n':
           line = line.strip("\n")
        if re.search(r'\.[0-9]+\.', line):
           line = line.strip(re.search(r'\.[0-9]+\.', line).group(0))
        file2.write(line)
    file1.close()
    file2.close()
    return path2

def date(stamp):
    """输出时间格式
    """
    delta = pd.Timedelta(str(stamp)+'D')
    real_time = pd.to_datetime('1899-12-30') + delta
    return real_time

def word_extraction_shanghai(province):
    i_out = 1
    path = r'E:\Spyderworkplace\excel\总实验数据'
    path_ = r'E:\Spyderworkplace\excel\总实验数据'
    filelist = os.listdir(path)
    '''
    for i in range(len(filelist)):
        if province in filelist[i]:
            path = os.path.join(path, filelist[i])
    # print(path)                                                 #找到省份文件夹——如“广西（齐全）”
    filelist = os.listdir(path)
    '''
    """ 生成excel的表头
    """
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
    # print(filelist)
    for file in filelist:
        if '.txt' in file and '2017' in file:
            word_path = os.path.join(path, file)
            document = open(word_path,encoding='utf-8')  # 打开docx
            name_shiwusuo = []  # 存事务所名称
            name_fuzeren = []  # 存总/分所所长（如有）
            print(word_path)
            year = re.search(r'\d{4}', file).group()  # 提取'年度'
            wb_read = xlrd.open_workbook(os.path.join(path_, '源地址.xlsx'))  # 从源地址excel提取'通知发布日期' '来源公告名称'
            for i in range(wb_read.sheet_by_index(0).nrows - 1):
                if year in wb_read.sheet_by_index(0).cell_value(i, 0) and province in wb_read.sheet_by_index(
                        0).cell_value(i, 0):
                    d = wb_read.sheet_by_index(0).cell(i, 1).value
                    date_pb = str(date(d)).split(" ")[0]
                    name_pb = wb_read.sheet_by_index(0).cell_value(i, 0)
    
            flag = -10
            n_name_shiwusuo = 0
            i=-1
            paragraphs = document.readlines()

            for paragraph in paragraphs:

                i=i+1
                #print(paragraph)
                if '2018' in file or '2019' in file:
                    if (re.search(r'代管', paragraph)):
                        break
                    if '序号' in paragraph:
                        continue
                    if (re.search(r'事务', paragraph) or re.search(r'分所', paragraph)) :  # 进入事务所



                        if '2018' in file:
                            name_kuaijishi = paragraph.split()[-2]
                            if '是' in paragraph:
                                id_kuaijishi = paragraph.split()[-4]
                                name_shiwusuo = paragraph.split()[-5]
                            else:
                                id_kuaijishi = paragraph.split()[-3]
                                name_shiwusuo = paragraph.split()[-4]
                        else:
                            name_kuaijishi = paragraph.split()[-3]
                            id_kuaijishi = paragraph.split()[-4]
                            name_shiwusuo = paragraph.split()[-5]
                        worksheet.write(i_out, 0, year)
                        worksheet.write(i_out, 1, date_pb)
                        worksheet.write(i_out, 2, name_pb)
                        worksheet.write(i_out, 3, province)
                        worksheet.write(i_out, 4, name_shiwusuo)  # 事务所
                        if '总所' in name_shiwusuo:
                            worksheet.write(i_out, 5, '总所')
                        elif '分所' in name_shiwusuo or '分公司' in name_shiwusuo:
                            worksheet.write(i_out, 5, '分所')
                        else:
                            worksheet.write(i_out, 5, '总所')
                        worksheet.write(i_out, 6, name_fuzeren)  # 负责人
                        worksheet.write(i_out, 7, name_kuaijishi.replace('。', ''))
                        worksheet.write(i_out, 8, id_kuaijishi)
                        i_out = i_out + 1
                else:
                    if (re.search(r'个人', paragraph)) or (re.search(r'尾部',paragraph)) \
                            or (re.search(r'暂缓', paragraph)) or (re.search(r'其他', paragraph)) :  #
                        #print(paragraph)
                        break 

                    if (re.search(r'事务', paragraph) or re.search(r'分所', paragraph))  \
                         or (re.search(r'代管', paragraph))  and '市（' not in paragraph:  # 进入事务所
                        #print(paragraph.split()[0])
                        name_shiwusuo = paragraph.split()[0]
                        name_shiwusuo = re.sub('[a-zA-Z0-9’!"#$%&\'()*+,-./:;<=>?@，：。?★、…【】《》？“”‘’！[\\]^_`{|}~\s]+', "",name_shiwusuo)  # 事务所名称
                        name_shiwusuo = re.sub('人', '', name_shiwusuo)
                        name_shiwusuo = re.sub('（）', '', name_shiwusuo)
                        name_shiwusuo = name_shiwusuo.split()[0]

                        flag = i


                    if i == flag:  # 进入人名
                        n = 1  # n是人名分段后的行数
                        #print(paragraph)
                        # print(paragraph_nextrow)

                        paragraph_nextrow=paragraphs[i+1]
                        #print(paragraph_nextrow)
                        while ('事务' not in paragraph_nextrow) and ('个人' not in paragraph_nextrow) \
                                and ('代管' not in paragraph_nextrow) and ('分所' not in paragraph_nextrow) \
                                and ('其他' not in paragraph_nextrow) and ('尾部' not in paragraph_nextrow):
                            #print(len(paragraph_nextrow))

                            paragraph = paragraph + ' ' + paragraph_nextrow
                            #print(paragraph)


                            paragraph_nextrow = paragraphs[i+1+n]  # pdf转word后多行人名
                            n = n + 1
                        #print(paragraph)
                        flag_pass = -1  # flag_pass 用来解决两个双字姓名相连问题
                        #print(paragraph.split())
                        for j in range(len(paragraph.split())):
                            #print(paragraph.split()[j])
                            # print(paragraph_nextrow)


                            if (len(paragraph.split()[j]) == 3) :  # 姓名为三字
                                name_kuaijishi = paragraph.split()[j]

                            elif (len(paragraph.split()[j]) == 4) :  # 姓名为三字
                                name_kuaijishi = paragraph.split()[j]

                            elif (len(paragraph.split()[j]) == 2) :  # 姓名为三字
                                name_kuaijishi = paragraph.split()[j]

                            elif (len(paragraph.split()[j]) * len(paragraph.split()[j - 1]) == 1) and (j - 1 != flag_pass):  # 姓名为二字，中间有空格
                                flag_pass = j
                                name_kuaijishi = paragraph.split()[j - 1] + paragraph.split()[j]




                            else:
                                continue


                            worksheet.write(i_out, 0, year)
                            worksheet.write(i_out, 1, date_pb)
                            worksheet.write(i_out, 2, name_pb)
                            worksheet.write(i_out, 3, province)
                            worksheet.write(i_out, 4, name_shiwusuo)  # 事务所
                            if '总所' in name_shiwusuo:
                                worksheet.write(i_out, 5, '总所')
                            elif '分所' in name_shiwusuo or '分公司' in name_shiwusuo:
                                worksheet.write(i_out, 5, '分所')
                            else:
                                worksheet.write(i_out, 5, '总所')
                            worksheet.write(i_out, 6, name_fuzeren)  # 负责人
                            worksheet.write(i_out, 7, name_kuaijishi.replace('。', ''))
                            #worksheet.write(i_out, 8, id_kuaijishi)
                            i_out = i_out + 1

    workbook.save(os.path.join(path_output,province+'.xls'))

if __name__ == "__main__":
    word_extraction_shanghai("上海2017")