# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: http://doc.scrapy.org/topics/item-pipeline.html

from scrapy.xlib.pydispatch import dispatcher
from scrapy import signals
from scrapy.contrib.exporter import CsvItemExporter
from scrapy.exceptions import DropItem

import time

from scrapy.xlib.pydispatch import dispatcher
from scrapy import signals
from scrapy.exceptions import DropItem

class CsvExportPipeline(object):

    def __init__(self):
        #self.duplicates = {}
        dispatcher.connect(self.spider_opened, signals.spider_opened)
        dispatcher.connect(self.spider_closed, signals.spider_closed)
        self.files = {}

    def spider_opened(self, spider):
        #self.duplicates[spider] = set()
        file = open('%s_%s.csv' % (spider.name, int(time.time())), 'w+b')
        self.files[spider] = file
        self.exporter = CsvItemExporter(file,fields_to_export = ['description','phone'])
        self.exporter.start_exporting()

    def spider_closed(self, spider):
        #del self.duplicates[spider]
        self.exporter.finish_exporting()
        file = self.files.pop(spider)
        file.close()

    def process_item(self, item, spider):
        #if item['description'] in self.duplicates[spider]:
        #    raise DropItem("Duplicateitemfound: %s" % item)
        #else:
        #    self.duplicates[spider].add(item['description'])
        #    self.exporter.export_item(item)
        #    return item
        if item is None:
            raise DropItem("None")
        self.exporter.export_item(item)
        return item 
			