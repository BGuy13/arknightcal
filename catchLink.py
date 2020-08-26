# !/usr/bin/python
# coding:utf-8

import requests as rq
from bs4 import BeautifulSoup
import io
import time
import mysql.connector
from opencc import OpenCC
cc = OpenCC('s2twp')

mydb = mysql.connector.connect(
    host = "192.168.168.146",
    user = "arkdeve",
    password = "BGuy6013@",
    database = "arknightcal"
)

fp = io.open("link01.csv", "w")
for i in range(1, 11):
    link = "https://www.diopoo.com/ark/characters?pn=" + str(i)
    print(link)
    # fp.write(link + "\n")
    nl_response = rq.get(link)
    soup = BeautifulSoup(nl_response.text, "html.parser")
    for character in soup.findAll('a', {'class': 'name'}):
        # print(cc.convert(character.string) + " -> https://www.diopoo.com/ark/" + character.get('href'))
        fp.write(cc.convert(character.text) + ",https://www.diopoo.com/ark/" + character.get('href') + "\n")
print("end")
fp.close()