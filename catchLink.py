# !/usr/bin/python
# coding:utf-8

import requests as rq
from bs4 import BeautifulSoup
import io
import time
import mysql.connector
from opencc import OpenCC
cc = OpenCC('s2twp')

# mydb = mysql.connector.connect(
#     host = "192.168.168.146",
#     user = "arkdeve",
#     password = "BGuy6013@",
#     database = "arknightcal"
# )

cnt = 0
for i in range(1, 2):
    link = "https://www.diopoo.com/ark/characters?pn=" + str(i)
    print(link)
    nl_response = rq.get(link)
    soup = BeautifulSoup(nl_response.text, "html.parser")
    for charSet in soup.findAll('div', {'class': 'round shadow'}):
        charJobStr = charSet.find('img', {'class': 'job shadow'}).get('src')
        charJobStr = charJobStr.split("/")[4].split(".")[0]
        charStarStr = charSet.find('img', {'class': 'star'}).get('src')
        charStarStr = charStarStr.split("_")[1].split(".")[0]
        print(str(cnt).zfill(3) + " : " + charJobStr + " -> " + charStarStr)
        cnt += 1
    # for charJob in soup.findAll('img', {'class': 'job shadow'}):
    #     print(charJob.get('src'))
    # for charStar in soup.findAll('img', {'class': 'star'}):
    #     print(charStar.get('src'))
    






# fp = io.open("link01.csv", "w")
# for i in range(1, 11):
#     link = "https://www.diopoo.com/ark/characters?pn=" + str(i)
#     print(link)
#     # fp.write(link + "\n")
#     nl_response = rq.get(link)
#     soup = BeautifulSoup(nl_response.text, "html.parser")
#     for character in soup.findAll('a', {'class': 'name'}):
#         # print(cc.convert(character.string) + " -> https://www.diopoo.com/ark/" + character.get('href'))
#         fp.write(cc.convert(character.text) + ",https://www.diopoo.com/ark/" + character.get('href') + "\n")
# print("end")
# fp.close()