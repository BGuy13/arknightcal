# !/usr/bin/python
# coding:utf-8

import requests as rq
from lxml import etree

headers = {"User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36"}
res = rq.get("https://www.diopoo.com/ark/character/723870",headers = headers)
content = res.content.decode("utf-8")
# print(content)
# print(res)

html = etree.HTML(content)

# 技能等級一到七
sk0 = {}
cnt = 1
for i in range(2, 8):
    sk0[cnt] = {}
    sk0[cnt]['level'] = i
    sk0[cnt]['materials'] = {}
    for j in range(1, 4):
        try:
            sk0[cnt]['materials'][j] = {}
            sk0[cnt]['materials'][j]['name'] = html.xpath('/html/body/div[2]/div[1]/div[2]/table[5]/tbody/tr[' + str(i-1) + ']/td/a[' + str(j) + ']/div/table/tr[1]/td[2]/text()')[0]
            sk0[cnt]['materials'][j]['quantity'] = html.xpath('/html/body/div[2]/div[1]/div[2]/table[5]/tbody/tr[' + str(i-1) + ']/td/a[' + str(j) + ']/span/text()')[0]
        except:
            del sk0[cnt]['materials'][j]
            break
    cnt += 1
# print(sk0)

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
                sk8_10[cnto][cnti]['materials'][j]['name'] = html.xpath('/html/body/div[2]/div[1]/div[2]/table[' + str(s+1) + ']/tbody/tr[5]/td[' + str(i-7) + ']/a[' + str(j) + ']/div/table/tr[1]/td[2]/text()')
                sk8_10[cnto][cnti]['materials'][j]['quantity'] = html.xpath('/html/body/div[2]/div[1]/div[2]/table[' + str(s+1) + ']/tbody/tr[5]/td[' + str(i-7) + ']/a[' + str(j) + ']/span/text()')
            except:
                del sk8_10[cnto][cnti]['materials'][j]
                break
        cnti += 1
    cnto += 1
# print(sk8_10)