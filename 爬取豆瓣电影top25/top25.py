# -*- coding: utf-8 -*-
# @Time     : 2021/4/18 16:01
# @File     : top25.py
# @Author   ：Rambo


from bs4 import BeautifulSoup
import requests
import xlwt


def analyse(url, datelist):
    header ={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36"
    }
    resp = requests.get(url, headers=header)
    soup = BeautifulSoup(resp.text, "html.parser")
    for item in soup.find_all('div', class_='item'):
        datelist.append(item.em.string)
        datelist.append(item.a.attrs["href"])
        datelist.append(item.span.string)


def save(datelist):
    k = -1
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('豆瓣250')
    for i in range(0,250):
        for j in range(0,3):
            k = k+1
            sheet.write(i, j, datelist[k])

    book.save('豆瓣.xls')


def main():
    datelist = []
    for i in range(0,10):
        url = "https://movie.douban.com/top250?start=" + str(i*25)
        analyse(url, datelist)

    save(datelist)

main()