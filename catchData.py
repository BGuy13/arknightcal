# !/usr/bin/python
# coding:utf-8

import requests as rq
from bs4 import BeautifulSoup
import io
from opencc import OpenCC
cc = OpenCC('s2twp')

fp = io.open("data01.csv", "w")
for i in range(1, 11):
    link = "https://www.diopoo.com/ark/items?pn=" + str(i)
    print(link)
    # fp.write(link + "\n")
    nl_response = rq.get(link)
    soup = BeautifulSoup(nl_response.text, "html.parser")
    for material in soup.findAll('b'):
        fp.write(cc.convert(material.text) + "\n")
print("end")
fp.close()