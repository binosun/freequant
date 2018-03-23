
# coding: utf-8

import urllib2
import urllib
import json
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

def urllib_requester(url,header):

    request = urllib2.Request(url, headers=header)
    response = urllib2.urlopen(request)
    html = response.read()
    if isinstance(html, str):
        try:
            url_response = html.decode('UTF-8')
            return url_response
        except:
            url_response = unicode(html)
            return url_response

if __name__ == "__main__":
    urllib_requester("https://xueqiu.com/service/partials/home/timeline?source=news&_=1512965940245")

