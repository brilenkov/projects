# Scrapy settings for suntimes project
#
# For simplicity, this file contains only the most important settings by
# default. All the other settings are documented here:
#
#     http://doc.scrapy.org/topics/settings.html
#

BOT_NAME = 'suntimes'
BOT_VERSION = '1.0'

SPIDER_MODULES = ['suntimes.spiders']
NEWSPIDER_MODULE = 'suntimes.spiders'
DEFAULT_ITEM_CLASS = 'suntimes.items.SunItem'
USER_AGENT = '%s/%s' % (BOT_NAME, BOT_VERSION)

ITEM_PIPELINES = [
    'suntimes.pipelines.CsvExportPipeline'
]

CONCURRENT_REQUESTS_PER_DOMAIN = 100
CONCURRENT_REQUESTS = 100
RANDOMIZE_DOWNLOAD_DELAY = False
DOWNLOAD_DELAY = 0
#LOG_LEVEL = 'WARNING'
LOG_LEVEL = 'DEBUG'
