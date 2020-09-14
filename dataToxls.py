# !/usr/bin/python
# coding:utf-8

import openpyxl
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import colors, Font, Fill, NamedStyle
from openpyxl.styles import PatternFill, Border, Side, Alignment
import json
import mysql.connector


def takeThird(elem):
    return elem[2]

def convert(cnt):
    cnt -= 65
    if cnt > 25:
        cntC = chr(64 + cnt // 26) + chr(65 + (cnt % 26))
    else:
        cntC = chr(65 + (cnt % 26))
    return cntC

arkdb = mysql.connector.connect(
    host="192.168.168.146",
    user="arkdeve",
    password="BGuy6013@",
    database="arknightDB"
)
cursor = arkdb.cursor()

cursor.execute("SELECT char_id, job, star, charName FROM charSummary")
datas = cursor.fetchall()

sorted = []
category = ["pioneer", "sniper", "medic", "caster", "warrior", "tank", "support", "special"]
Job = ["先鋒", "狙擊", "醫療", "術師", "近衛", "重裝", "輔助", "特種"]
for i in range(0, 8):
    temp = []
    for data in datas:
        if data[1] == category[i]:
            temp.append(data)
    temp.sort(key=takeThird, reverse=True)
    sorted.append(temp)
# print(sorted)

wb = Workbook()

wc = wb['Sheet']
wc.title = 'Char'
wd = wb.create_sheet('Data')

for j in range(0,8):
    wd.merge_cells(convert(65 + j*4) + '1:' + convert(68 + j*4) + '1')
    # print(convert(65 + j*4) + '1:' + convert(68 + j*4) + '1')
    wd[convert(65 + j*4) + '1'].value = Job[j]
    wd[convert(65 + j*4) + '1'].alignment = Alignment(horizontal='center', vertical='center')
    cnt = 2
    for char in sorted[j]:
        for i in range(0, 4):
            wd[convert(65 + j*4 + i) + str(cnt)].value = char[i]
            # print(convert(65 + j*4 + i) + str(cnt))
            wd[convert(65 + j*4 + i) + str(cnt)].alignment = Alignment(horizontal='center', vertical='center')
        cnt += 1

wb.save('Test.xlsx')