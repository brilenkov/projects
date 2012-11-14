from urlparse import urljoin
from scrapy import log
from scrapy.http import HtmlResponse
from scrapy.utils.response import get_meta_refresh
from scrapy.exceptions import IgnoreRequest, NotConfigured
from scrapy.conf import settings
from twisted.internet.error import TimeoutError as ServerTimeoutError, DNSLookupError, \
                                   ConnectionRefusedError, ConnectionDone, ConnectError, \
                                   ConnectionLost, TCPTimedOutError
from twisted.internet.defer import TimeoutError as UserTimeoutError
import base64
import random

from settings import USER_AGENT
import socket
import urllib2
import lxml.html
import datetime
import psycopg2

DB_NAME = 'postgres'
DB_USER = 'postgres'
DB_PSWD = '123456'

conn2 = psycopg2.connect('dbname=' + DB_NAME + ' user=' + DB_USER + ' password=' + DB_PSWD)
cur2 = conn2.cursor()

#-------------------------------------------------------
#                       CheckItemDB
#-------------------------------------------------------
def CheckItemDB(ID, CHK_ITEM, Item):
    cur2.execute(CHK_ITEM, (str(ID),))
    Item = cur2.fetchone()
    if Item is None:
        return True
    return False

#-------------------------------------------------------
#                       DUPLICATES
#-------------------------------------------------------
class IgnoreDupMiddleware(object):
    def process_request(self, request, spider):
        CHK_ITEM = "SELECT * FROM cl_urls WHERE url = %s;"  
        try:
           ID = request.url
        except:
           ID = ''
        Item = None
        if ID == '' or CheckItemDB(ID, str(CHK_ITEM), Item):
            if ID == '': ID = request.url
            log.msg('Detected unique request(%s)' % ID, log.DEBUG, spider=spider)
            return
        else:
            if ID == '': ID = request.url
            log.msg('Detected duplicate request(ID=%s)' % ID, log.DEBUG, spider=spider)
            if Item is not None:
                if request.url != Item[-1]:
                    log.msg('Two ID with different URL',    log.ERROR, spider=spider)
                    log.msg('request.url %s' % request.url, log.ERROR, spider=spider)
                    log.msg('Item[-1]    %s' % Item[-1],    log.ERROR, spider=spider)
            raise IgnoreRequest
'''
# This middleware can be used to avoid re-visiting already visited items, which can be useful for speeding up the scraping for projects with immutable items, ie. items that, once scraped, don't change.

from scrapy import log
from scrapy.http import Request
from scrapy.item import BaseItem
from scrapy.utils.request import request_fingerprint

from myproject.items import MyItem

class IgnoreVisitedItems(object):
    """Middleware to ignore re-visiting item pages if they were already visited
    before. The requests to be filtered by have a meta['filter_visited'] flag
    enabled and optionally define an id to use for identifying them, which
    defaults the request fingerprint, although you'd want to use the item id,
    if you already have it beforehand to make it more robust.
    """

    FILTER_VISITED = 'filter_visited'
    VISITED_ID = 'visited_id'
    CONTEXT_KEY = 'visited_ids'

    def process_spider_output(self, response, result, spider):
        context = getattr(spider, 'context', {})
        visited_ids = context.setdefault(self.CONTEXT_KEY, {})
        ret = []
        for x in result:
            visited = False
            if isinstance(x, Request):
                if self.FILTER_VISITED in x.meta:
                    visit_id = self._visited_id(x)
                    if visit_id in visited_ids:
                        log.msg("Ignoring already visited: %s" % x.url,
                                level=log.INFO, spider=spider)
                        visited = True
            elif isinstance(x, BaseItem):
                visit_id = self._visited_id(response.request)
                if visit_id:
                    visited_ids[visit_id] = True
                    x['visit_id'] = visit_id
                    x['visit_status'] = 'new'
            if visited:
                ret.append(MyItem(visit_id=visit_id, visit_status='old'))
            else:
                ret.append(x)
        return ret

    def _visited_id(self, request):
        return request.meta.get(self.VISITED_ID) or request_fingerprint(request)

# Snippet imported from snippets.scrapy.org (which no longer works)
# author: pablo
# date  : Aug 10, 2010
'''