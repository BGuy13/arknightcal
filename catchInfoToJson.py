# !/usr/bin/python
# coding:utf-8

import requests as rq
from lxml import etree
import io
import json
import mysql.connector
from opencc import OpenCC
cc = OpenCC('s2t')

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
fp = io.open("check.json", "w", encoding="utf-8")
for data in datas:
    res = rq.get(data[2], headers = headers)
    content = res.content.decode("utf-8")

    info[data[0]] = {}
    info[data[0]]["name"] = data[1]
    info[data[0]]["demand"] = {}
    de = info[data[0]]["demand"]

    de["skill"] = {}
    sk = de["skill"]

    de["elite"] = {}
    el = de["elite"]

    html = etree.HTML(content)
    print("start")

    for i in range(2, 8):
        sk["0_" + str(i)] = {}
        sk["0_" + str(i)]["materials"] = {}
        ski = sk["0_" + str(i)]["materials"]
        for j in range(1, 4):
            try:
                ski["m" + str(j)] = {}
                ski["m" + str(j)]["name"] = cc.convert(html.xpath("/html/body/div[2]/div[1]/div[2]/table[5]/tbody/tr[" + str(i-1) + "]/td/a[" + str(j) + "]/div/table/tr[1]/td[2]/text()")[0])
                ski["m" + str(j)]["quantity"] = html.xpath("/html/body/div[2]/div[1]/div[2]/table[5]/tbody/tr[" + str(i-1) + "]/td/a[" + str(j) + "]/span/text()")[0]
            except:
                del ski["m" + str(j)]
                break
    # print(sk)

    for s in range(1, 4):
        for i in range(8, 11):
            sk[str(s) + "_" + str(i)] = {}
            sk[str(s) + "_" + str(i)]["materials"] = {}
            sksi = sk[str(s) + "_" + str(i)]["materials"]
            for j in range(1, 4):
                try:
                    sksi["m" + str(j)] = {}
                    sksi["m" + str(j)]["name"] = cc.convert(html.xpath("/html/body/div[2]/div[1]/div[2]/table[" + str(s+1) + "]/tbody/tr[5]/td[" + str(i-7) + "]/a[" + str(j) + "]/div/table/tr[1]/td[2]/text()")[0])
                    sksi["m" + str(j)]["quantity"] = html.xpath("/html/body/div[2]/div[1]/div[2]/table[" + str(s+1) + "]/tbody/tr[5]/td[" + str(i-7) + "]/a[" + str(j) + "]/span/text()")[0]
                except:
                    del sksi["m" + str(j)]
                    break
    
    for i in range(1, 3):
        el[str(i)] = {}
        el[str(i)]["materials"] = {}
        eli = el[str(i)]["materials"]
        for j in range(1, 6):
            try:
                eli["m" + str(j)] = {}
                eli["m" + str(j)]["name"] = cc.convert(html.xpath("/html/body/div[2]/div[1]/div[2]/table[" + str(i+5) + "]/tbody/tr[4]/td/a[" + str(j) + "]/div/table/tr[1]/td[2]/text()")[0])
                if not eli["m" + str(j)]["name"]:
                    del eli["m" + str(j)]
                    break
                eli["m" + str(j)]["quantity"] = html.xpath("/html/body/div[2]/div[1]/div[2]/table[" + str(i+5) + "]/tbody/tr[4]/td/a[" + str(j) + "]/span/text()")[0]
            except:
                del eli["m" + str(j)]
                break
        
json.dump(info, fp, ensure_ascii=False)
fp.close()