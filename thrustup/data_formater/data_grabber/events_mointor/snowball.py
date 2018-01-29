# -*-coding=utf-8-*-
#抓取雪球的收藏文章
__author__ = 'Rocky'
import cookielib
import json
import re
import requests
import time
from lxml import etree

from freequant.thrustup.data_formater.common.toolkit import Toolkit

url='https://xueqiu.com/snowman/login'
session = requests.session()

session.cookies = cookielib.LWPCookieJar(filename="cookies")
try:
    session.cookies.load(ignore_discard=True)
except:
    print "Cookie can't load"

agent = 'Mozilla/5.0 (Windows NT 5.1; rv:33.0) Gecko/20100101 Firefox/33.0'
headers = {'Host': 'xueqiu.com',
           'Referer': 'https://xueqiu.com/',
           'Origin':'https://xueqiu.com',
           'User-Agent': agent}
account=Toolkit.getUserData('data.cfg')
print account['snowball_user']
print account['snowball_password']

data={'username':account['snowball_user'],'password':account['snowball_password']}
s=session.post(url,data=data,headers=headers)
print s.status_code
#print s.text
session.cookies.save()
fav_temp="https://xueqiu.com/statuses/stock_timeline.json?symbol_id=SH600276&count=10&source=公告&page=1"
collection=session.get(fav_temp,headers=headers)
print "collection",collection.text
fav_content= collection.text
p=re.compile('"maxPage":(\d+)')
print p.findall(fav_content)
maxPage=p.findall(fav_content)[0]
print maxPage
print type(maxPage)
maxPage=int(maxPage)
print type(maxPage)
for i in range(1,maxPage+1):
    fav='https://xueqiu.com/favs?page=%d' %i
    collection=session.get(fav,headers=headers)
    fav_content= collection.text
    #print fav_content
    p=re.compile('var favs = {(.*?)};',re.S|re.M)
    result=p.findall(fav_content)[0].strip()

    new_result='{'+result+'}'
    #print type(new_result)
    #print new_result
    data=json.loads(new_result)
    use_data= data['list']
    host='https://xueqiu.com'
    for i in use_data:
        url=host+ i['target']
        print url
        txt_content=session.get(url,headers=headers).text
        #print txt_content.text

        tree=etree.HTML(txt_content)
        title=tree.xpath('//title/text()')[0]

        filename = re.sub('[\/:*?"<>|]', '-', title)
        print filename

        content=tree.xpath('//div[@class="detail"]')
        Toolkit.save2filecn(filename,"Link: %s\n\n" %url)
        for i in content:
            Toolkit.save2filecn(filename, i.xpath('string(.)'))
        #print content
        #Toolkit.save2file(filename,)
        time.sleep(10)
