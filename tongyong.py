import pandas as pd
import numpy as np
import datetime
import os
import re

# forVoucherNo 用于循环赋予 excel 们凭证号（排名不分先后）
# forItemNo 用于给每一个 excel 内部存在的科目


def CrossOver(dir, type, pl, fl):
    for i in os.listdir(dir):  # 遍历整个文件夹
        path = os.path.join(dir, i)
        # print(os.path.splitext(path)[1])
        if os.path.isfile(path) and os.path.splitext(path)[1] == type:  # 判断是否为一个文件，排除文件夹
            pl.append(path)
            fl.append(i)
        elif os.path.isdir(path):
            newdir = path
            CrossOver(newdir, type, pl, fl)
    return pl, fl


def forVoucherNo(datadir, exportDir):
    # 递归获取这里面的所有子孙目录下的 xlsx 后缀，然后循环赋予凭证号
    if os.path.exists(exportDir) == False:
        os.mkdir(exportDir)
    pathlist = []
    namelist = []
    willprocess_excels_path, willprocess_excels_name = CrossOver(dir=datadir, type='.xlsx', pl=pathlist, fl=namelist)
    print(len(willprocess_excels_path))
    print(len(willprocess_excels_name))
    noRecord = 0  # 计数器，用于记下没有数据生成的凭证确保一致
    for index in range(len(willprocess_excels_name)):
        df1 = pd.read_excel(willprocess_excels_path[index])
        if len(df1) == 0:
            noRecord = noRecord + 1  # 用于去掉没有数据的
            continue
        df1['凭证号'] = index + 1 - noRecord
        # print(f'{exportDir}'+'\\'+f'{willprocess_excels_name[index]}')
        print(index + 1 - noRecord)
        df1.to_excel(f'{exportDir}'+'\\'+f'{willprocess_excels_name[index]}',index=False)
    print(len(willprocess_excels_name) - noRecord)


def forItemNo(datadir, exportDir):
    # 递归获取这里面的所有子孙目录下的 xlsx 后缀，分别处理分录号
    if os.path.exists(exportDir) == False:
        os.mkdir(exportDir)
    pathlist = []
    namelist = []
    willprocess_excels_path, willprocess_excels_name = CrossOver(dir=datadir, type='.xlsx', pl=pathlist, fl=namelist)
    for index in range(len(willprocess_excels_name)):
        df1 = pd.read_excel(willprocess_excels_path[index])
        # print(df1['摘要'].tolist())
        zhaiyao = df1['摘要'].tolist()
        # print(zhaiyao)
        tempList = list(set(zhaiyao))
        tempList.sort(key=zhaiyao.index)
        for i in range(len(tempList)):
            df1.loc[df1['摘要'] == tempList[i], '分录号'] = i + 1
        # print(f'{exportDir}'+'\\'+f'{willprocess_excels_name[index]}')
        df1.to_excel(f'{exportDir}'+'\\'+f'{willprocess_excels_name[index]}',index=False)


if __name__ == "__main__":
    forVoucherNo(datadir=r'.\pingzheng', exportDir=r'.\pingzheng_1')
    forItemNo(datadir=r'.\pingzheng_1', exportDir=r'.\pingzheng_2')

