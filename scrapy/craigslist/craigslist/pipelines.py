#!/usr/bin/env python
# encoding=utf-8
# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: http://doc.scrapy.org/topics/item-pipeline.html

from scrapy.xlib.pydispatch import dispatcher
from scrapy import log, signals
from scrapy.contrib.exporter import CsvItemExporter
from scrapy.exceptions import DropItem

import time
import datetime
import traceback

#----------------------------------------------------------------------------#
# CSV Exporter                                                               #
#----------------------------------------------------------------------------#
class CsvExportPipeline(object):

    def __init__(self):
        dispatcher.connect(self.spider_opened, signals.spider_opened)
        dispatcher.connect(self.spider_closed, signals.spider_closed)
        self.files = {}

    def spider_opened(self, spider):
        file_other = open('%s_%s.csv' % (spider.name + '_allother', int(time.time())), 'w+b')
        file_room = open('%s_%s.csv' % (spider.name +  '_roomates', int(time.time())), 'w+b')
        self.files[spider] = file_other
        self.files[spider] = file_room
        self.exporter_other = CsvItemExporter(file_other,fields_to_export = 
        [
        'url',
        'listingid',
        'email',
        'city',
        'state',
        'country',
        'province',
        'address',
        'location',
        'date',
        'title',
        'details',
        'userbody'
        ])
        self.exporter_room = CsvItemExporter(file_room,fields_to_export = 
        [
        'url',
        'listingid',
        'email',
        'city',
        'state',
        'country',
        'province',
        'address',
        'location',
        'date',
        'title',
        'details',
        'userbody'
        ])

    def spider_closed(self, spider):
        self.exporter_other.finish_exporting()
        self.exporter_room.finish_exporting()
        file_other = self.files.pop(spider)
        file_other.close()
        file_room = self.files.pop(spider)
        file_room.close()

    def process_item(self, item, spider):
        if item is None:
            raise DropItem("None")
            
        if '/roo/' in str(item['url']):
            self.exporter_room.export_item(item)
        else:
            self.exporter_other.export_item(item)
        return item

def convstr(s):
    try:
        if s is None:
            out = ''
        elif str(s).strip() == '':
            out = ''
        else:
            try:
                out = ''.join(s)
            except:
                try:
                    out = str(s)
                except:
                    out = ''
    except:
        try:
            out = s.encode('utf-8')
        except:
            out = ''
    return out
    
import psycopg2
DB_NAME = 'postgres'
DB_USER = 'postgres'
DB_PSWD = '123456'
#CREATE TABLE cl_urls( url text PRIMARY KEY)
#CREATE TABLE cl_allother( id serial, url text PRIMARY KEY, listingid text, email text, city text, state text, country text, province text, address text, location text, date text, title text, details text, userbody text, adddate timestamp without time zone)
#CREATE TABLE cl_roomates( id serial, url text PRIMARY KEY, listingid text, email text, city text, state text, country text, province text, address text, location text, date text, title text, details text, userbody text, adddate timestamp without time zone)
#----------------------------------------------------------------------------#
# DataBase Exporter                                                          #
#----------------------------------------------------------------------------#
class DBExportPipeline(object):

    def __init__(self):
        dispatcher.connect(self.spider_opened, signals.spider_opened)
        dispatcher.connect(self.spider_closed, signals.spider_closed)
    
    
    def spider_opened(self, spider):
        self.conn = psycopg2.connect('dbname=' + DB_NAME + ' user=' + DB_USER + ' password=' + DB_PSWD)
        #self.conn2 = psycopg2.connect('dbname=' + DB_NAME + ' user=' + DB_USER + ' password=' + DB_PSWD)
        #self.conn3 = psycopg2.connect('dbname=' + DB_NAME + ' user=' + DB_USER + ' password=' + DB_PSWD)
        self.cur = self.conn.cursor()
        #self.cur2 = self.conn2.cursor()
        #self.cur3 = self.conn3.cursor()
        #self.ADD_ITEM3 = "INSERT INTO cl_urls (url) VALUES (%s);"
        self.ADD_ITEM1 = "INSERT INTO cl_allother (url, listingid, email, city, state, country, province, address, location, date, title, details, userbody, adddate) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);"
        self.ADD_ITEM2 = "INSERT INTO cl_roomates (url, listingid, email, city, state, country, province, address, location, date, title, details, userbody, adddate) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);"
        self.ADD_ITEM3 = "INSERT INTO cl_urls (url) VALUES (%s);"
        self.CHK_ITEM1 = "SELECT * FROM cl_allother WHERE email = %s;"
        self.CHK_ITEM2 = "SELECT * FROM cl_roomates WHERE email = %s;"

    def spider_closed(self, spider):
        self.cur.close()
        #self.cur2.close()
        #self.cur3.close()
        self.conn.close()
        #self.conn2.close()
        #self.conn3.close()
    def process_item(self, item, spider):
        self.cur.execute(self.ADD_ITEM3,(str(item['url']),))
        self.conn.commit()
        if item['email'] is None or str(item['email']).strip() == '':
            raise DropItem("None")
            return True
            
        if item is None:
            raise DropItem("None")
        log.msg("Store Item " + item['url'], log.DEBUG, spider=spider)
        
        if self.CheckItemDB(item['email']):
            #log.msg("CheckItemDB branch " + item['url'] , log.DEBUG, spider=spider)
            #try:
            if '/roo/' in str(item['url']):
                    #log.msg("Store Item finally roo" + item['url'], log.DEBUG, spider=spider)
                    self.cur.execute(self.ADD_ITEM2,
                                      (str(item['url']), 
                                       str(item['listingid']),
                                       str(item['email']),
                                       str(item['city']),    
                                       str(item['state']),
                                       str(item['country']),   
                                       convstr(item['province']),
                                       convstr(item['address']),   
                                       convstr(item['location']),
                                       convstr(item['date']),
                                       convstr(item['title']), 
                                       convstr(item['details']), 
                                       convstr(item['userbody']),
                                       datetime.datetime.now(),))
            else:
                    log.msg("Store Item finally " + item['url'], log.DEBUG, spider=spider)
                    self.cur.execute(self.ADD_ITEM1,
                                      (str(item['url']), 
                                       str(item['listingid']),
                                       str(item['email']),
                                       str(item['city']),    
                                       str(item['state']),
                                       str(item['country']),   
                                       convstr(item['province']),
                                       convstr(item['address']),   
                                       convstr(item['location']),
                                       convstr(item['date']),
                                       convstr(item['title']), 
                                       convstr(item['details']), 
                                       convstr(item['userbody']),
                                       datetime.datetime.now(),))
                    
            #except:
            #    log.msg("Unable to add Item " + item['url'], log.ERROR, spider=spider)
            #    log.msg(traceback.format_exc(), log.ERROR)
                
            self.conn.commit()
            #self.conn2.commit()
            #self.conn3.commit()
            '''
            try:
                self.cur1.execute(self.ADD_ITEM3,
                                  (item['url'],))
            except:
                log.msg("Unable to add Item " + item['url'], log.ERROR, spider=spider)
                log.msg(traceback.format_exc(), log.ERROR)
            self.conn1.commit()
            '''
        return item

    def CheckItemDB(self, ID):
        if '/roo/' in ID:
            self.cur.execute(self.CHK_ITEM2, (str(ID),))
        else:
            self.cur.execute(self.CHK_ITEM1, (str(ID),))
        if self.cur.fetchone() is None:
            
            return True
        log.msg("Item " + ID + " is already in db", log.ERROR, spider='cl')
        return False