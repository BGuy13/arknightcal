# !/usr/bin/python
# coding:utf-8

import requests as rq
from lxml import etree
import io
import mysql.connector
from opencc import OpenCC
cc = OpenCC('s2twp')

arkdb = mysql.connector.connect(
    host = "192.168.168.146",
    user = "arkdeve",
    password = "BGuy6013@",
    database = "arknightDB"
)
cursor = arkdb.cursor()

cursor.execute("SELECT char_id, charName, infoLink FROM charSummary")
datas = cursor.fetchall()
headers = {"User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36"}
info = {}
cntO = 0
fp = io.open("check.txt", "w", encoding="utf-8")
for data in datas:
    res = rq.get(data[2], headers = headers)
    content = res.content.decode("utf-8")
    sk0 = {}
    sk0['char_id'] = data[0]
    sk0['charName'] = data[1]
    # print(str(data[0]) + ":" + data[1])
    # fp.write(data[1] + "\n")

    html = etree.HTML(content)
    print('start')

    # 技能等級一到七
    cnt = 1
    for i in range(2, 8):
        sk0[cnt] = {}
        sk0[cnt]['level'] = i
        sk0[cnt]['materials'] = {}
        for j in range(1, 4):
            try:
                sk0[cnt]['materials'][j] = {}
                sk0[cnt]['materials'][j]['name'] = cc.convert(html.xpath('/html/body/div[2]/div[1]/div[2]/table[5]/tbody/tr[' + str(i-1) + ']/td/a[' + str(j) + ']/div/table/tr[1]/td[2]/text()')[0])
                sk0[cnt]['materials'][j]['quantity'] = html.xpath('/html/body/div[2]/div[1]/div[2]/table[5]/tbody/tr[' + str(i-1) + ']/td/a[' + str(j) + ']/span/text()')[0]
            except:
                del sk0[cnt]['materials'][j]
                break
        cnt += 1
    # print(sk0)
    fp.write(str(sk0) + "\n")

    # 技能等級八到十
    sk8_10 = {}
    cnto = 1
    for s in range(1, 4):
        sk8_10[cnto] = {}
        cnti = 7
        for i in range(8, 11):
            sk8_10[cnto][cnti] = {}
            sk8_10[cnto][cnti]['level'] = i
            sk8_10[cnto][cnti]['materials'] = {}
            for j in range(1, 4):
                try:
                    sk8_10[cnto][cnti]['materials'][j] = {}
                    sk8_10[cnto][cnti]['materials'][j]['name'] = cc.convert(html.xpath('/html/body/div[2]/div[1]/div[2]/table[' + str(s+1) + ']/tbody/tr[5]/td[' + str(i-7) + ']/a[' + str(j) + ']/div/table/tr[1]/td[2]/text()')[0])
                    sk8_10[cnto][cnti]['materials'][j]['quantity'] = html.xpath('/html/body/div[2]/div[1]/div[2]/table[' + str(s+1) + ']/tbody/tr[5]/td[' + str(i-7) + ']/a[' + str(j) + ']/span/text()')[0]
                except:
                    del sk8_10[cnto][cnti]['materials'][j]
                    break
            cnti += 1
        cnto += 1
    # print(sk8_10)
    fp.write(str(sk8_10) + "\n")

    # 菁英化
    elite = {}
    cnt = 1
    for i in range(1, 3):
        elite[cnt] = {}
        elite[cnt]['level'] = i
        elite[cnt]['materials'] = {}
        for j in range(1, 6):
            try:
                elite[cnt]['materials'][j] = {}
                elite[cnt]['materials'][j]['name'] = cc.convert(html.xpath('/html/body/div[2]/div[1]/div[2]/table[' + str(i+5) + ']/tbody/tr[4]/td/a[' + str(j) + ']/div/table/tr[1]/td[2]/text()')[0])
                if not elite[cnt]['materials'][j]['name']:
                    # print("in")
                    del elite[cnt]['materials'][j]
                    break
                elite[cnt]['materials'][j]['quantity'] = html.xpath('/html/body/div[2]/div[1]/div[2]/table[' + str(i+5) + ']/tbody/tr[4]/td/a[' + str(j) + ']/span/text()')
            except:
                del elite[cnt]['materials'][j]
                break
        cnt += 1
    # print(elite)
    fp.write(str(elite) + "\n")
    print('end')

print('everything_end')
fp.close()