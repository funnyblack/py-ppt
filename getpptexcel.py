from typing import Mapping
from xml.sax.handler import property_interning_dict
from pptx import Presentation
import os
import pandas as pd
import numpy as np

writer = pd.ExcelWriter(
    r'excelinfo.xlsx', engine='openpyxl')  # 写入excel表格
colNumIdx = 0
pptIdx = 0
files = os.listdir('/Users/leslie/py/py-ppt')
files.sort()
for i in files:
    if i.endswith(".pptx"):
        pptx = Presentation(i)
        slideMap = {}  # 创建一个map对象用来存slide内容，方便后续操作
        rowNumIdx = 0
        tableIdx = 0
        for slide in pptx.slides:  # 遍历ppt页
            for shape in slide.shapes:  # 遍历页中所有方框
                cell_list = []  # 储存ppt表格的所有单元格内容
                first_cell_list = []
                if shape.has_table:  # 判断方框中是否有表格
                    if tableIdx == 0:
                        if pptIdx == 0:
                            cell_list.append("")
                            cell_list.append(i)
                        else:
                            cell_list.append(i)
                    rownum = len(shape.table.rows)  # 获取表格的行
                    colnum = len(shape.table.columns)  # 获取表格的列
                    for row in range(rownum):
                        for coloum in range(colnum):
                            if pptIdx != 0 and coloum == 0:
                                continue
                            cell = shape.table.cell(
                                row, coloum).text  # 遍历后获取单元格内容
                            cell_list.append(cell)  # 将单元格内容存到列表中
                    if pptIdx != 0:
                        if tableIdx == 0:
                            table_base = np.array(cell_list).reshape(
                                rownum+1, colnum - 1)  # 用numpy构建数组
                        else:
                            table_base = np.array(cell_list).reshape(
                                rownum, colnum - 1)  # 用numpy构建数组
                    else:
                        if tableIdx == 0:
                            table_base = np.array(cell_list).reshape(
                                rownum+1, colnum)  # 用numpy构建数组
                        else:
                            table_base = np.array(cell_list).reshape(
                                rownum, colnum)  # 用numpy构建数组
                    table = pd.DataFrame(table_base)  # 用pandas构建二维数据
                    table.to_excel(
                        writer, startrow=rowNumIdx, startcol=colNumIdx, index=False, header=False)  # 存储到表格
                    tableIdx += 1
            rowNumIdx = rowNumIdx + rownum
        if pptIdx != 0:
            colNumIdx = colNumIdx + colnum - 1
        else:
            colNumIdx = colNumIdx + colnum
        pptIdx += 1
writer.save()
