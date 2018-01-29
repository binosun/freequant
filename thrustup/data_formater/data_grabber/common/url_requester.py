
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
    print "html", type(html), html
    # if isinstance(html, str):
    #     url_response = unicode(html)
    #     return url_response
    # else:
    url_response = html.decode('UTF-8')
    return url_response

if __name__ == "__main__":
    print "hello"
    urllib_requester("https://xueqiu.com/service/partials/home/timeline?source=news&_=1512965940245")

