#!/usr/bin/env python
# encoding: utf-8


"""
@version: ??
@author: phpergao
@license: Apache Licence 
@contact: endoffight@gmail.com
@site: http://
@software: PyCharm
@file: readCensusExcel.py
@time: 2016/3/15 21:59
"""


# readCensusExcel.py - Tabulates population and number of census tracts for
# each county.

import openpyxl, pprint

print('Opening workbook...')
wb =openpyxl.load_workbook('censuspopdata.xlsx')
sheet = wb.get_sheet_by_name('Population by Census Tract')
countyData = {}

print('Reading rows...')
for row in range(2, sheet.get_highest_row() + 1):
    state = sheet['B'+str(row)].value
    county = sheet['C'+str(row)].value
    pop = sheet['D'+str(row)].value
    # countyData[state abbrev][county]['tracts']
    # countyData[state abbrev][county]['pop']
    countyData.setdefault(state, {})
    countyData[state].setdefault(county, {'tracts':0, 'pop':0})
    countyData[state][county]['tracts'] += 1
    countyData[state][county]['pop'] += int(pop)

print('Writing results...')
resultFile = open('census2010.py', 'w')
resultFile.write('allData = ' + pprint.pformat(countyData))
resultFile.close()
print('Done.')
