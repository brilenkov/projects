#!/usr/bin/env python
# encoding=utf-8

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/topics/items.html

from scrapy.item import Item, Field

class CraigslistItem(Item):
    #listingid, city, state, country, province, etc, firstname and lastname
    listingid = Field()
    city = Field()
    state = Field()
    country = Field()
    province = Field()
    address = Field()
    location = Field()
    email = Field()
    url = Field()
    date = Field()
    title = Field()
    details = Field()
    userbody = Field()
