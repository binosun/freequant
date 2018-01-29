
#coding: utf-8

import cookielib
import requests

from freequant.thrustup.data_formater.common.toolkit import Toolkit


def make_session():
    url='https://xueqiu.com/snowman/login'
    session = requests.session()

    session.cookies = cookielib.LWPCookieJar(filename="cookies")
    try:
        session.cookies.load(ignore_discard=True)
    except:
        print "Cookie can't load"

    data, header = create_data_header()
    s=session.post(url,data=data,headers=header)
    print s.status_code
    session.cookies.save()

    return session, header

def create_data_header():
    agent = 'Mozilla/5.0 (Windows NT 5.1; rv:33.0) Gecko/20100101 Firefox/33.0'
    header = {'Host'      : 'xueqiu.com',
               'Referer'   : 'https://xueqiu.com/',
               'Origin'    : 'https://xueqiu.com',
               'User-Agent': agent}
    account = Toolkit.getUserData('config/user.cfg')
    data = {'username': account['snowball_user'], 'password': account['snowball_password']}
    return data, header