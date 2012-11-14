#!/usr/bin/env python
# encoding=utf-8
# Scrapy settings for craigslist project
#
# For simplicity, this file contains only the most important settings by
# default. All the other settings are documented here:
#
#     http://doc.scrapy.org/topics/settings.html
#

BOT_NAME = 'craigslist'
BOT_VERSION = '1.0'

SPIDER_MODULES = ['craigslist.spiders']
NEWSPIDER_MODULE = 'craigslist.spiders'
DEFAULT_ITEM_CLASS = 'craigslist.items.CraigslistItem'
USER_AGENT = 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1'

ITEM_PIPELINES = [
    #'craigslist.pipelines.CsvExportPipeline'
    'craigslist.pipelines.DBExportPipeline' 
]
DOWNLOADER_MIDDLEWARES = {
    'craigslist.middlewares.IgnoreDupMiddleware': 450,
}
RETRY_PRIORITY_ADJUST = +2
RETRY_TIMES = 5

CONCURRENT_REQUESTS_PER_DOMAIN = 100
CONCURRENT_REQUESTS = 100
#RANDOMIZE_DOWNLOAD_DELAY = True
#DOWNLOAD_DELAY = 0.2
#COOKIES_ENABLED = False
LOG_FILE   = 'craigslist.log'
LOG_STDOUT = True
#LOG_LEVEL = 'INFO'
#LOG_LEVEL = 'WARNING'
LOG_LEVEL = 'DEBUG'