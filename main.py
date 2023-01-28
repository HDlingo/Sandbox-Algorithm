import matplotlib.pyplot as plt
import numpy as np
import matplotlib
import xlrd  # 导入xlrd库

# dataset <List>  \\  len(dataset) 获取元素列表个数（长度）  \\list.count(元素) 查看元素在列表中出现的次数
file = 'D:/杂货/编码数据.xlsx'  # 文件路径
# 双文件导入，存入dataset,dataInfer
fileCont = './data/test.xls'  # 用该方法读取表格和表单里的单元格的数据
fileInfer = './data/factor.xls'  # 用该方法读取表格和表单里的单元格的数据

# --------全局变量区--------#
dictSandbox = {}  # 字典，用来统计各沙具的数量信息   eg: plane:2
dataset = []  # Sandbox信息汇总
dataInfer = []  # 沙具属性列表       eg：loneliness-Security：1
AttriDict = {}  # 字典，用来统计属性值  eg：loneliness-Security：1


# --------输出/索引函数--------#
def ListPrint(datalist, col, row):  # 列表输出函数
    for c in col:
        for r in row:
            print(datalist[c][r], end=" ")
        print('\n')
    return 0


def ListInfer(dataInfer, col, row):  # 列表输出函数
    for c in range(0, col):
        for r in range(0, row):
            print(dataInfer[c][r], end=" ")
        print('\n')
    return 0


def DictSandboxPrint():  # 字典输出函数
    for key, value in dictSandbox.items():  # 遍历输出字典中所有键值队
        print(key, value)


def SearchSandbox(index):
    return dataInfer[index][0]


# --------统计/处理函数--------#

def DictAmount(col):  # 沙具统计函数，存入字典 col为data.xls的信息列数
    for c in range(1, col):
        index = (int)(dataset[c][0])
        str = dataInfer[index][0]
        if dictSandbox.get(str) is None:  # 查询字典是否存在当前字段
            dictSandbox[str] = 1
        else:
            dictSandbox[str] += 1
    # DictSandboxPrint()


def ArrtibuteGet(row):  # 属性类别获取
    for c in range(1, row):
        AttriDict[dataInfer[0][c]] = 0;


def getElementPos(sandboxName):  # 查询sandbox类型在属性xls文件当中的位置，方便获取其各属性值
    for c in range(1, wordList_Infer.nrows):
        if dataInfer[c][0] == sandboxName:
            return c


def ArrtibuteAmount(row):  # 属性值统计 第一层基本统计
    # ListInfer(dataInfer, wordList_Infer.nrows, wordList_Infer.ncols)
    for key, value in dictSandbox.items():
        pos = getElementPos(key)
        for tpe in range(1, wordList_Infer.ncols):
            str=dataInfer[0][tpe]
            AttriDict[str]+=dataInfer[pos][tpe]*value    # value表示同一种沙具摆放的数量
    return 0
        # --------主代码区--------#

pageIndex = xlrd.open_workbook(filename=fileCont)  # 用方法打开该文件路径下的文件
wordList = pageIndex.sheet_by_name("Sheet1")  # 打开该表格里的表单

for r in range(wordList.nrows):  # 遍历行
    col = []
    for l in range(wordList.ncols):  # 遍历列
        col.append(wordList.cell(r, l).value)  # 将单元格中的值加入到列表中(r,l)相当于坐标系，cell（）为单元格，value为单元格的值
    dataset.append(col)

from pprint import pprint  # pprint的输出形式为一行输出一个结果，下一个结果换行输出。实质上pprint输出的结果更为完整

# pprint(dataset)    # 输出 列表中的元素
# print('\n')
# ListPrint(dataset,range(wordList.nrows),range(wordList.ncols))

pageIndex_Infer = xlrd.open_workbook(filename=fileInfer)  # 用方法打开该文件路径下的文件
wordList_Infer = pageIndex_Infer.sheet_by_name("Sheet1")  # 打开该表格里的表单

for r in range(wordList_Infer.nrows):
    col = []
    for l in range(wordList_Infer.ncols):
        col.append(wordList_Infer.cell(r, l).value)
    dataInfer.append(col)

DictAmount(wordList.nrows)  # 沙具统计函数，存入字典 col为data.xls的信息列数
ArrtibuteGet(wordList_Infer.ncols)  # 属性类别获取
ArrtibuteAmount(wordList.nrows)  # 属性值统计 第一层基本统计

for key, value in AttriDict.items():  # 遍历输出字典中所有键值队
    print(key, value)