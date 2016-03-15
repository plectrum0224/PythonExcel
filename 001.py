#!/usr/bin/env python
# encoding: utf-8


"""
@version: ??
@author: phpergao
@license: Apache Licence 
@contact: endoffight@gmail.com
@site: http://
@software: PyCharm
@file: 001.py
@time: 2016/3/15 20:21
"""
import openpyxl
from openpyxl.utils import *

wb = openpyxl.load_workbook('example.xlsx')

# get a list of all the sheet names in the workbook
sheet = wb.get_sheet_names()
# ['Sheet1', 'Sheet2', 'Sheet3']

sheet1 = wb.get_sheet_by_name('Sheet1')
# # ---------------------------------------------
# print(sheet1['A1'].value)
# # 2015-04-05 13:34:02
# print(sheet1['B1'].value)
# # Apples
# print(sheet1['B1'].row)
# # 1
# print(sheet1['B1'].column)
# # B
# print(sheet1['B1'].coordinate)
# # B1
# print(sheet1['C1'].value)
# # 73
# # ---------------------------------------------
#
# for i in range(1, 8, 2):
#     print(i, sheet1.cell(row=i, column=2).value)
# 1 Apples
# 3 Pears
# 5 Apples
# 7 Strawberries

# print(tuple(sheet1['A1':'C3']))
# ((<Cell Sheet1.A1>, <Cell Sheet1.B1>, <Cell Sheet1.C1>),
# (<Cell Sheet1.A2>,<Cell Sheet1.B2>, <Cell Sheet1.C2>),
# (<Cell Sheet1.A3>, <Cell Sheet1.B3>,<Cell Sheet1.C3>))
# for rowOfCellObjects in sheet1['A1':'C3']:
#     for cellObj in rowOfCellObjects:
#         print(cellObj.coordinate, cellObj.value)
#     print('---END OF ROW---')
# A1 2015-04-05 13:34:02
# B1 Apples
# C1 73
# ---END OF ROW---
# A2 2015-04-05 03:41:23
# B2 Cherries
# C2 85
# ---END OF ROW---
# A3 2015-04-06 12:46:51
# B3 Pears
# C3 14
# ---END OF ROW---

# print(sheet1.columns[1])
# (<Cell Sheet1.B1>, <Cell Sheet1.B2>, <Cell Sheet1.B3>,
# <Cell Sheet1.B4>, <Cell Sheet1.B5>, <Cell Sheet1.B6>,
# <Cell Sheet1.B7>)
for cellObj in sheet1.columns[1]:
    print(cellObj.value)
# Apples
# Cherries
# Pears
# Oranges
# Apples
# Bananas
# Strawberries

