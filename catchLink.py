# !/usr/bin/python
# coding:utf-8

import requests as rq
from bs4 import BeautifulSoup
import io
# import mysql.connector
from opencc import OpenCC
cc = OpenCC('s2twp')

# arkdb = mysql.connector.connect(
#     host = "192.168.168.146",
#     user = "arkdeve",
#     password = "BGuy6013@",
#     database = "arknightDB"
# )
# cursor = arkdb.cursor()

char = {}
cnt = 0
for i in range(1, 11):
    link = "https://www.diopoo.com/ark/characters?pn=" + str(i)
    print(link)
    nl_response = rq.get(link)
    soup = BeautifulSoup(nl_response.text, "html.parser")
    tmpCnt = cnt
    for charSet1 in soup.findAll('div', {'class': 'round shadow'}):
        char[cnt] = {}
        charJobStr = charSet1.find('img', {'class': 'job shadow'}).get('src')
        charJobStr = charJobStr.split("/")[4].split(".")[0]
        charStarStr = charSet1.find('img', {'class': 'star'}).get('src')
        charStarStr = charStarStr.split("_")[1].split(".")[0]
        char[cnt]["id"] = str(cnt + 1)
        char[cnt]["job"] = charJobStr
        char[cnt]["star"] = str(int(charStarStr) + 1)
        cnt += 1
    for charSet2 in soup.findAll('a', {'class': 'name'}):
        char[tmpCnt]["name"] = cc.convert(charSet2.string)
        char[tmpCnt]["link"] = "https://www.diopoo.com/ark/" + charSet2.get('href')
        tmpCnt += 1

fp = io.open("charList.csv", "w", encoding="utf-8")
fp.write("char_id,job,star,charName,infoLink\n")
for j in range(0, cnt):
    if j == cnt - 1:
        fp.write(char[j]['id'] + "," + char[j]['job'] + "," + char[j]['star'] + "," + char[j]['name'] + "," + char[j]['link'])
    else:
        fp.write(char[j]['id'] + "," + char[j]['job'] + "," + char[j]['star'] + "," + char[j]['name'] + "," + char[j]['link'] + "\n")
print("end")
fp.close()