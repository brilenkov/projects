from scrapy.contrib.spiders import CrawlSpider, Rule
from scrapy.contrib.linkextractors.sgml import SgmlLinkExtractor
from scrapy.selector import HtmlXPathSelector
from scrapy.item import Item
import urllib2
import re
from craigslist.items import CraigslistItem
from decimal import Decimal
import unicodedata
from BeautifulSoup import BeautifulSoup
import lxml.html
from scrapy.spider import BaseSpider

from scrapy.http.response.html import HtmlResponse
from scrapy.http import Request 
from scrapy import log

CITIES_STATES = {}
CITIES_COUNTRIES = {}
STATES_COUNTRIES = {}
            
class CraigslistSpider(CrawlSpider):
    name = "cl"
    
    allowed_domains = [
    "craigslist.org",  
    "craigslist.ca",
    "craigslist.at",
    "craigslist.cz",
    "craigslist.fi",
    "craigslist.de",
    "craigslist.gr",
    "craigslist.pl",
    "craigslist.pt",
    "craigslist.es",
    "craigslist.co.uk",
    "craigslist.se",
    "craigslist.it",
    "craigslist.ch",
    "craigslist.tr",
    "craigslist.com.cn",
    "craigslist.hk",
    "craigslist.co.in",
    "craigslist.jp",
    "craigslist.co.kr",
    "craigslist.com.ph",
    "craigslist.com.sg",
    "craigslist.com.tw",
    "craigslist.com.th",
    "craigslist.com.au",
    "craigslist.co.nz",
    "craigslist.com.mx",
    "craigslist.co.za"
    ]
    start_urls = ['http://www.craigslist.org/about/sites']
    rules = (
        Rule(SgmlLinkExtractor(restrict_xpaths=('//*[@class="colmask"]/div/div/div/div/ul/li/a'))),
        Rule(SgmlLinkExtractor(restrict_xpaths=('//*[@class="sublinks"]/a'))),
        Rule(SgmlLinkExtractor(allow=('.*/appartments/.*','.*/apa/.*','.*/roo/.*','.*/sub/.*','.*/hsw/.*','.*/swp/.*','.*/vac/.*','.*/prk/.*','.*/off/.*','.*/rea/.*',))),
        Rule(SgmlLinkExtractor(restrict_xpaths=('//*[@id="nextpage"]/font/a'))),
        Rule(SgmlLinkExtractor(restrict_xpaths=('//*[@class="row"]/a',)),callback='parse_item'),
    )                                           
    def parse_start_url(self, response):
        hxs = HtmlXPathSelector(response)
        colmasks = hxs.select('//*[@class="colmask"]')
        colmaskcount = 0
        for colmask in colmasks[:-1]:
            colmaskcount+=1
            h1 = ''.join(hxs.select('//*[@class="colmask"][' + str(colmaskcount) + ']/h1/text()').extract())
            statecount = 0
            states = hxs.select('//*[@class="colmask"][' + str(colmaskcount) + ']/div/div/div/div/div/text()').extract()
            for state in states:
                statecount +=1
                STATES_COUNTRIES[state] = h1
                cities = hxs.select('//*[@class="colmask"][' + str(colmaskcount) + ']/div/div/div/div/ul[' + str(statecount) + ']/li/a').extract()
                for city in cities:
                    city = city[city.find('://')+3:city.find('.')]
                    CITIES_STATES[city] = state
                    CITIES_COUNTRIES[city] = h1

    def parse_item(self, response):
        
        item = CraigslistItem() 
        email_pattern = re.compile(r'''(?:[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?\.)+[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[A-Za-z0-9-]*[A-Za-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])''',re.IGNORECASE)
        soup = BeautifulSoup(response.body)
        tree = lxml.html.document_fromstring(response.body)
        #print 'referer:' + response.request.headers.get('Referer')
        hxs = HtmlXPathSelector(response)
        vref = response.request.headers.get('Referer')
        #a = email_pattern.findall(user['description'])
        a = email_pattern.findall(str(soup))
        a = set(a)
        try:
            if 'craigslist' not in ''.join(a):
                item['email'] = ','.join(a)
            else:
                item['email'] = None
        except:
            item['email'] = None
        
        item['url'] = response.url
        item['listingid'] = response.url[response.url.rfind('/')+1:-5]
        
        try:
            maps = tree.xpath('//*[@id="userbody"]/small/a')
            for map in maps:
                if 'google map' in map.text: 
                    themap = map.attrib['href'] 
                    adr = themap.split('+')
                    break
                elif 'yahoo map' in map.text:
                    themap = map.attrib['href'].replace('&country=', '+').replace('&csz=','+').replace('?addr=','+')
                    adr = themap.split('+')
                    break
        except:
            pass
            
        try:
            item['address'] = ' '.join(adr[1:-3]).replace('%3','').replace('q=', '').replace('+',' ').replace('%2','.')
        except:
            item['address'] = None
        
        try:
            item['date'] = ''.join(hxs.select('//*[@class="postingdate"]/text()').extract()).replace('Date: ','')
        except:
            item['date'] = None
            
        try:
            item['title'] = hxs.select('/html/body/h2/text()').extract()
        except:
            item['title'] = None
        try:
            item['details'] = re.sub('\s+',' ',re.sub('<[^<>]+>|&[a-z]+;',' ',' '.join(hxs.select('//*[@class="blurbs"]').extract()))).strip()
        except:
            item['details'] = None
        try:
            item['userbody'] = re.sub('\s+',' ',re.sub('<[^<>]+>|&[a-z]+;',' ',' '.join(hxs.select('//*[@id="userbody"]').extract()))).strip()
        except:
            item['userbody'] = None

        try:
            item['city'] = response.url[response.url.find('://')+3:response.url.find('.')]
        except:
            item['city'] = None
        
        try:
            item['country'] = CITIES_COUNTRIES[item['city']]
        except:
            item['country'] = None
        try:
            item['state'] = CITIES_STATES[item['city']]
        except:
            item['state'] = None
        
        try:
            locationtags = hxs.select('//*[@class="blurbs"]/li').extract()
            for locationtag in locationtags:
                if 'Location: ' in re.sub('\s+',' ',re.sub('<[^<>]+>|&[a-z]+;',' ',locationtag)).strip(): 
                    item['location'] = re.sub('\s+',' ',re.sub('<[^<>]+>|&[a-z]+;',' ',locationtag)).strip().replace('Location: ', '')
                    break
                else:
                    item['location'] = None
        except:
            item['location'] = None
        
        if len(response.url.split('/')) == 6:
            try:
                item['province'] = response.url.split('/')[3]
            except:
                item['province'] = None
        else:
            item['province'] = None
            
        #if item['email'] is None or item['email'] == '':
        #    item = None    
        return item 
