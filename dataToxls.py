# !/usr/bin/python
# coding:utf-8

from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import colors, Font, Fill, NamedStyle
from openpyxl.styles import PatternFill, Border, Side, Alignment
import json
import mysql.connector

def takeThird(elem):
    return elem[2]

arkdb = mysql.connector.connect(
    host = "192.168.168.146",
    user = "arkdeve",
    password = "BGuy6013@",
    database = "arknightDB"
)
cursor = arkdb.cursor()

cursor.execute("SELECT char_id, job, star, charName FROM charSummary")
datas = cursor.fetchall()

sorted = []
category = ["pioneer", "sniper", "medic", "caster", "warrior", "tank", "support", "special"]
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

wd.merge_cells('A1:D1')
wd['A1'].value = '先鋒'
wd['A1'].alignment = Alignment(horizontal='center', vertical='center')
cnt = 2
for char in sorted[0]:
    for i in range(0, 4):
        wd[chr(65 + i) + str(cnt)].value = char[i]
        wd[chr(65 + i) + str(cnt)].alignment = Alignment(horizontal='center', vertical='center')
    cnt += 1

# print(chr(65+1) + '{}')

wb.save('Test.xlsx')

# wb = Workbook()

# ws = wb.create_sheet('New Sheet')

# wa = wb['Sheet']

# for number in range(1,100): #Generates 99 "ip" address in the Column A;
#     ws['A{}'.format(number)].value= "192.168.1.{}".format(number)

# data_val = DataValidation(type="list", formula1="='New Sheet'!$A$1:$A$99") #You can change =$A:$A with a smaller range like =A1:A9
# wa.add_data_validation(data_val)

# data_val.add(wa["B1"]) #If you go to the cell B1 you will find a drop down list with all the values from the column A

# wb.save('Test.xlsx')