#! /usr/bin/python
#! -*- coding: utf-8 -*-

from scrapy.contrib.spiders import CrawlSpider, Rule
from scrapy.contrib.linkextractors.sgml import SgmlLinkExtractor
from scrapy.selector import HtmlXPathSelector
from scrapy.item import Item
import re
from suntimes.items import SunItem

from BeautifulSoup import BeautifulSoup
from decimal import Decimal

import codecs
encoding = 'cp1251'
codec_info = codecs.lookup(encoding)

f = codecs.StreamReaderWriter(open('phones.txt', 'U'), codec_info[2], codec_info[3], 'strict')
lines = f.readlines()
ln = 0
phones = []
for line in lines:
    phones.append(line)
f.close()

class SuntimesSpider(CrawlSpider):
    name = "suntimes"
    allowed_domains = ["http://suntimes.ru", "suntimes.ru", "www.suntimes.ru", "http://www.suntimes.ru"]
    start_urls = [
        "http://suntimes.ru/tickets-5.html",
        "http://suntimes.ru/tickets-6.html",
        "http://suntimes.ru/tickets-7.html",
        "http://suntimes.ru/tickets-8.html",
        "http://suntimes.ru/tickets-9.html",
        "http://suntimes.ru/tickets-11.html",
        "http://suntimes.ru/tickets-12.html",
        "http://suntimes.ru/tickets-13.html",
        "http://suntimes.ru/tickets-14.html",
        "http://suntimes.ru/tickets-16.html",
        "http://suntimes.ru/tickets-17.html",
        "http://suntimes.ru/tickets-18.html",
        "http://suntimes.ru/tickets-19.html",
        "http://suntimes.ru/tickets-20.html",
        "http://suntimes.ru/tickets-22.html",
        "http://suntimes.ru/tickets-23.html",
        "http://suntimes.ru/tickets-24.html",
        "http://suntimes.ru/tickets-25.html",
        "http://suntimes.ru/tickets-26.html",
        "http://suntimes.ru/tickets-28.html",
        "http://suntimes.ru/tickets-29.html",
        "http://suntimes.ru/tickets-30.html",
        "http://suntimes.ru/tickets-231.html",
        "http://suntimes.ru/tickets-232.html",
        "http://suntimes.ru/tickets-32.html",
        "http://suntimes.ru/tickets-33.html",
        "http://suntimes.ru/tickets-34.html",
        "http://suntimes.ru/tickets-35.html",
        "http://suntimes.ru/tickets-36.html",
        "http://suntimes.ru/tickets-37.html",
        "http://suntimes.ru/tickets-38.html"
    ]
    
    rules = (
        #Rule(SgmlLinkExtractor(restrict_xpaths=('//*[@id="subtopick"]/a', '//*[@id="Ssubtopick"]/a','//*[@id="text"]/a',)),callback='parse_item'),
        #Rule(SgmlLinkExtractor(restrict_xpaths=('//*[@id="Ssubtopick"]/a','//*[@id="text"]/a',)),callback='parse_item'),
        Rule(SgmlLinkExtractor(restrict_xpaths=('//*[@id="text"]/a',)),callback='parse_item'),
    )
    
    def parse_item(self, response):
        g = codecs.open('desc.txt','a', encoding='utf-8')
        g_without = codecs.open('desc_without.txt','a', encoding='utf-8')
        
        item = SunItem()

        soup = BeautifulSoup(response.body)
        items = []
        
        lis = soup.find('td',{'id': 'text'}).find('ul').findAll('li')
        
        for li in lis:
            try:
                item['description'] = li.text.strip()
            except:
                item['description'] = None
            try:
                
                agency = False

                mbphones = re.findall('\s([0-9]{1}[-][0-9]{3}[-][0-9]{3}[-][0-9]{2}[-][0-9]{2})', li.text.strip())
                for mb in mbphones:
                    if mb in phones:
                        agency = True
                        
                stphones = re.findall('\s([0-9]{2}[-][0-9]{2}[-][0-9]{2})', li.text.strip())
                for mb in stphones:
                    if mb in phones:
                        agency = True
                        
                item['phone'] = ','.join(mbphones) + ',' + ','.join(stphones)
            except:
                item['phone'] = None
            
                
            if not agency:
                if item['description']:
                    g_without.write(item['description'])
                    if item['phone']:
                        g_without.write('|') 
                        g_without.write(item['phone']) 
                    g_without.write('\n')
                    
            if item['description']:
                g.write(item['description'])
                if item['phone']:
                    g.write('|') 
                    g.write(item['phone']) 
                g.write('\n')

            #if item['description'] is None and item['phone'] is None:
            #    item = None
            items.append(item)
        
        g.close()
        g_without.close()
        return items
