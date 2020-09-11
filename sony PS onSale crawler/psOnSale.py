import requests
from bs4 import BeautifulSoup
import json
import tablib
import time

def sleeptime(hour,min,sec):
    return hour*3600 + min*60 + sec

def combineUrl (url,page):
    return str.format(url, page) 

def getSonyOnSale(url):
    page = 1
    formatedUrl = combineUrl(url,page)
    try:
        res = requests.get(formatedUrl)
        while(formatedUrl == res.url):
            soup = BeautifulSoup(res.text, 'html.parser')

            ItemList = soup.findAll('div',class_="grid-cell__body")

            for item in ItemList:
                yield {
                '遊戲名稱' : item.find('span').get('title'),
                '價錢' : item.h3.text,
                '連結' : "https://store.playstation.com" + item.find('a').get('href'),
                }
            page = page + 1
            formatedUrl = combineUrl(url,page)
            time.sleep(sleeptime(0,0,5))
            res = requests.get(formatedUrl)
    except:
        print("An exception occurred")
        
def start():
    url = input("請輸入特價網址:")
    url_fornt = url[:url.index('?')]
    url_end = url[url.index('?'):]
    url_fornt = url_fornt.split('/')
    url_fornt[-1] = "{0}"
    url_processed = '/'.join(url_fornt) + url_end
    items = getSonyOnSale(url_processed)
    s = json.dumps(list(items), indent = 4, ensure_ascii=False)
    rows = json.loads(s)
   # 将json中的key作为header, 也可以自定义header（列名）
    header=tuple([ i for i in rows[0].keys()])

    data = []
# 循环里面的字典，将value作为数据写入进去
    for row in rows:
        body = []
        for v in row.values():
            body.append(v)
        data.append(tuple(body))

    data = tablib.Dataset(*data,headers=header)

    open('psOnSale.xlsx', 'wb').write(data.export('xlsx'))

start()

