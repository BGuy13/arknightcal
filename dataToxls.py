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
import io

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

cursor.execute("SELECT char_id FROM charSummary ORDER BY char_id DESC")
lastID = cursor.fetchone()
lastID = lastID[0]
# print(lastID[0])

sorted = []
category1 = ["pioneer", "sniper", "medic", "caster", "warrior", "tank", "support", "special"]
category2 = ["0_2", "0_3", "0_4", "0_5", "0_6", "0_7", "1_8", "1_9", "1_10", "2_8", "2_9", "2_10", "3_8", "3_9", "3_10"]
category3 = [1, 3, 2, 3, 2, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3]
category4 = ["m1", "m2", "m3", "m4"]
Job = ["先鋒", "狙擊", "醫療", "術師", "近衛", "重裝", "輔助", "特種"]
for i in range(0, 8):
    temp = []
    for data in datas:
        if data[1] == category1[i]:
            temp.append(data)
    temp.sort(key=takeThird, reverse=True)
    sorted.append(temp)
# print(sorted)

f = io.open("check.json", "r", encoding="utf-8")
checks = json.load(f)

wb = Workbook()

wc = wb['Sheet']
wc.title = 'Char'
ws = wb.create_sheet('Summary')
wd = wb.create_sheet('Skill')
we = wb.create_sheet('Elite')

# 角色簡介
for j in range(0,8):
    # ws.merge_cells(convert(65 + j*3) + '1:' + convert(67 + j*3) + '1')
    # ws[convert(65 + j*3) + '1'].value = Job[j]
    # ws[convert(65 + j*3) + '1'].alignment = Alignment(horizontal='center', vertical='center')
    cnt = 2
    star = 6
    temp = 0
    ws[convert(65 + j*3 + 2) + '1'].value = "----6----"
    ws[convert(65 + j*3 + 2) + '1'].alignment = Alignment(horizontal='center', vertical='center')
    for char in sorted[j]:
        for i in range(0, 3):
            if char[2] != star:
                star = char[2]
                ws[convert(65 + j*3 + 2) + str(cnt + temp)].value = "----" + str(star) + "----"
                ws[convert(65 + j*3 + 2) + str(cnt + temp)].alignment = Alignment(horizontal='center', vertical='center')
                temp += 1
            if i > 1:
                ws[convert(65 + j*3 + i) + str(cnt + temp)].value = char[i+1]
                ws[convert(65 + j*3 + i) + str(cnt + temp)].alignment = Alignment(horizontal='center', vertical='center')
            else:
                ws[convert(65 + j*3 + i) + str(cnt + temp)].value = char[i]
                ws[convert(65 + j*3 + i) + str(cnt + temp)].alignment = Alignment(horizontal='center', vertical='center')
            if i%3 == 0:
                ws.column_dimensions[convert(65 + j*3 + i)].width = 5
            elif i%3 == 1:
                ws.column_dimensions[convert(65 + j*3 + i)].width = 7
            else:
                ws.column_dimensions[convert(65 + j*3 + i)].width = 11
        cnt += 1

# 技能需求
temp = 66
for i in range(0, 15):
    wd.merge_cells(convert(temp) + '1:' + convert(temp + category3[i]*2 - 1) + '1')
    wd[convert(temp) + '1'].value = 'S' + category2[i]
    wd[convert(temp) + '1'].alignment = Alignment(horizontal='center', vertical='center')
    for j in range(0, category3[i]):
        wd.merge_cells(convert(temp + j*2) + '2:' + convert(temp + j*2 + 1) + '2')
        wd[convert(temp + j*2) + '2'].value = category4[j]
        wd[convert(temp + j*2) + '2'].alignment = Alignment(horizontal='center', vertical='center')
    temp += category3[i]*2

for i in range(1, lastID):
    check = checks[str(i)]
    wd['A' + str(i+2)].value = check['name']
    wd['A' + str(i+2)].alignment = Alignment(horizontal='center', vertical='center')
    sk = check['demand']['skill']
    temp = 66
    for j in range(0, 15):
        for k in range(0, category3[j]):
            try:
                wd[convert(temp + k*2) + str(i+2)].value = sk[category2[j]]['materials'][category4[k]]['name']
                wd[convert(temp + k*2) + str(i+2)].alignment = Alignment(horizontal='center', vertical='center')
                wd.column_dimensions[convert(temp + k*2)].width = 13
                wd[convert(temp + k*2 + 1) + str(i+2)].value = sk[category2[j]]['materials'][category4[k]]['quantity']
                wd[convert(temp + k*2 + 1) + str(i+2)].alignment = Alignment(horizontal='center', vertical='center')
                wd.column_dimensions[convert(temp + k*2 + 1)].width = 5
            except:
                break
        temp += category3[j]*2
wd.column_dimensions['A'].width = 11

# 菁英化需求
for i in range(0, 2):
    we.merge_cells(convert(66 + i*8) + '1:' + convert(66 + i*8 + 7) + '1')
    we[convert(66 + i*8) + '1'].value = 'E' + str(i+1)
    we[convert(66 + i*8) + '1'].alignment = Alignment(horizontal='center', vertical='center')
    for j in range(0, 4):
        we.merge_cells(convert((66 + i*8) + j*2) + '2:' + convert((66 + i*8) + j*2 + 1) + '2')
        we[convert((66 + i*8) + j*2) + '2'].value = category4[j]
        we[convert((66 + i*8) + j*2) + '2'].alignment = Alignment(horizontal='center', vertical='center')

for i in range(1, lastID):
    check = checks[str(i)]
    we['A' + str(i+2)].value = check['name']
    we['A' + str(i+2)].alignment = Alignment(horizontal='center', vertical='center')
    el = check['demand']['elite']
    for j in range(0, 2):
        for k in range(0, 4):
            try:
                we[convert(66 + j*8 + k*2) + str(i+2)].value = el[str(j+1)]['materials'][category4[k]]['name']
                we[convert(66 + j*8 + k*2) + str(i+2)].alignment = Alignment(horizontal='center', vertical='center')
                we.column_dimensions[convert(66 + j*8 + k*2)].width = 13
                we[convert((66 + j*8 + k*2) + 1) + str(i+2)].value = el[str(j+1)]['materials'][category4[k]]['quantity']
                we[convert((66 + j*8 + k*2) + 1) + str(i+2)].alignment = Alignment(horizontal='center', vertical='center')
                if k%4 == 0:
                    we.column_dimensions[convert((66 + j*8 + k*2) + 1)].width = 7
                else:
                    we.column_dimensions[convert((66 + j*8 + k*2) + 1)].width = 4
            except:
                break
we.column_dimensions['A'].width = 11

wb.save('Test.xlsx')
f.close()