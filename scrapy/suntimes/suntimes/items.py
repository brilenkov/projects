# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/topics/items.html

from scrapy.item import Item, Field

class SunItem(Item):
    # define the fields for your item here like:
    description = Field()
    phone = Field()

class Website(SunItem):

    url = Field()

    def __str__(self):
        return "Website: name=%s url=%s" % (self.get('description'), self.get('url'))
