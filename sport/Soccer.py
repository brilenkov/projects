#! /usr/bin/python
#! -*- coding: utf-8 -*-
import Queue
import threading
import urllib, urllib2
import time, re
import optparse                          
from xlwt import *
import codecs
from BeautifulSoup import BeautifulSoup
import datetime
from datetime import date
from punky import Browster
import os
from openpyxl import load_workbook
from win32com.client import Dispatch
import sys
reload(sys) 
sys.setdefaultencoding('utf-8') 

debugF = False
#debugF = True
firstlegues = [
'/soccer/argentina/primera-division/',
'/soccer/austria/t-mobile-bundesliga/',
'/soccer/belgium/jupiler-league/',
'/soccer/brazil/campeonato-brasileiro/',
'/soccer/chile/primera-division/',
'/soccer/colombia/liga-postobon/',
'/soccer/croatia/1-hnl/',
'/soccer/czech-republic/gambrinus-liga/',
'/soccer/denmark/superliga/',
'/soccer/england/premier-league/',
'/soccer/finland/veikkausliiga/',
'/soccer/france/ligue-1/',
'/soccer/germany/bundesliga/',
'/soccer/greece/super-league/',
'/soccer/hungary/nb-i/',
'/soccer/iceland/landsbankadeild/',
'/soccer/ireland/premier-league/',
'/soccer/italy/serie-a/',
'/soccer/japan/j-league/',
'/soccer/mexico/primera-division/',
'/soccer/netherlands/eredivisie/',
'/soccer/norway/tippeligaen/',
'/soccer/paraguay/primera-division/',
'/soccer/peru/primera-division/',
'/soccer/poland/ekstraklasa/',
'/soccer/portugal/portuguese-liga/',
'/soccer/romania/liga-i/',
'/soccer/russia/premier-league/',
'/soccer/scotland/premier-league/',
'/soccer/serbia/super-liga/',
'/soccer/slovakia/corgon-liga/',
'/soccer/slovenia/prva-liga/',
'/soccer/spain/primera-division/',
'/soccer/sweden/allsvenskan/',
'/soccer/switzerland/super-league/',
'/soccer/turkey/superliga/',
'/soccer/ukraine/premier-league/',
'/soccer/uruguay/primera-division/',
'/soccer/usa/major-league-soccer/'
]
def weekDay():
    out = ''
    if curDate.weekday() == 0:
        out = 'M'
    elif curDate.weekday() == 1:
        out = 'TU'
    elif curDate.weekday() == 2:
        out = 'W'
    elif curDate.weekday() == 3:
        out = 'TH'
    elif curDate.weekday() == 4:
        out = 'F'
    elif curDate.weekday() == 5:
        out = 'SA'
    elif curDate.weekday() == 6:
        out = 'SU'
    return out

queue = Queue.Queue()

class ThreadUrl(threading.Thread):
    """Threaded Url Grab"""
    def __init__(self, queue):
        threading.Thread.__init__(self)
        self.queue = queue

    def run(self):
        while True:
            host = self.queue.get()
            #vname = host[host.rfind('=')+1:] + '.htm'
            vname = host.replace(':','^').replace('/','_').replace('?','+') + '.htm'
            if os.path.exists('SoccerPastRes/' + vname):
                vd = os.path.getmtime('SoccerPastRes/' + vname)
                vdd = datetime.date.fromtimestamp(vd)
                if vdd < curDate:
                    urllib.urlretrieve(host, 'SoccerPastRes/' + vname)
            else:
                urllib.urlretrieve(host, 'SoccerPastRes/' + vname)
            self.queue.task_done()

class ThreadUrlRead(threading.Thread):
    """Threaded Url Grab"""
    def __init__(self, queue):
        threading.Thread.__init__(self)
        self.queue = queue

    def run(self):
        while True:
            vhost = self.queue.get()
            attempts = 0
            while attempts < 10: 
                try:
                    html = urllib2.urlopen(vhost, None)
                    attempts += 1
                    break
                except urllib2.HTTPError, error: 
                    print "*** WARNING: Reconnecting after 4 sec..."
                    html = ''
                    time.sleep(4)
                    attempts += 1
                    
                print '*** ERROR: Unable to connect after ' + str(attempts) + ' tries!'
            soup = BeautifulSoup(''.join(html))
            legTableLinks = soup.find('div', attrs={'id':'main'}).find('div', attrs={'id':'ml'}).findAll('a')
            if legTableLinks:
                for legTableLink in legTableLinks:
                    if legTableLink['href'] in firstlegues:
                        print 'DEBUG: ThreadUrlRead: ' + 'http://www.betexplorer.com' + legTableLink['href'] + tz
                        firstLinks.append('http://www.betexplorer.com' + legTableLink['href'] + tz)
                        break
                else:
                    print '*** WARNING: ThreadUrlRead: ' + vhost + ' doesnot contain first legues'
            else:
                print '*** WARNING: ThreadUrlRead: ' + vhost + ' doesnot contain legues'
            self.queue.task_done()

class ThreadUrlReadGames(threading.Thread):
    """Threaded Url Grab"""
    def __init__(self, queue):
        threading.Thread.__init__(self)
        self.queue = queue

    def run(self):
        while True:
            vhost = self.queue.get()
            attempts = 0
            while attempts < 10: 
                try:
                    html = urllib2.urlopen(vhost, None)
                    attempts += 1
                    break
                except urllib2.HTTPError, error: 
                    print "*** WARNING: Connection ERROR. Reconnecting after 4 sec..."
                    html = ''
                    time.sleep(4)
                    attempts += 1
                    
                print '*** ERROR: Unable to connect after ' + str(attempts) + ' tries!'
            
            soup = BeautifulSoup(''.join(html))
            GameLinksTable = soup.find('table', attrs={'id':'league-summary-next'})
            #print GameLinksTable
            rows = GameLinksTable.findAll('tr') # get all rows in this table
            #print rows
            for tr in rows: 
                #print tr
                cols = tr.findAll('td') # find all columns in table
                
                for td in cols[1:2]: 
                    #print td
                    if td.find('a'):
                        gameDateText = cols[7].find(text=True)
                        #print gameDateText
                        gameDateArr = gameDateText.split(' ')[0].split('.')
                        gameDate = date(int(gameDateArr[2]), int(gameDateArr[1]), int(gameDateArr[0]))
                        #print gameDate
                        #print curDate
                        if gameDate == curDate:
                            #I am only looking for the games with a coefficient above 2.00 for both teams!!!
                            #In other words, if any of the teams has a coefficient lower than 2.00 that 
                            #game should be skipped.
                            #try:
                            #    print cols[0].text.encode('utf-8') + ' / ' + str(float(cols[3].text.encode('utf-8'))) + ':' + str(float(cols[5].text.encode('utf-8')))
                            #except:
                                #print 'debug ' + '='*40
                            #    pass
                            #print float(cols[4].find(text=True))
                            #print float(cols[6].find(text=True))
                            if float(cols[4].find(text=True)) >= 2.00 and float(cols[6].find(text=True)) >= 2.00:
                                GameLinks.append(vhost[:vhost.rfind('/')+1] + td.find('a')['href'] + '&' + tz[1:])
                                try:
                                    print 'DEBUG: ' + cols[1].text.encode('utf-8') + ' added'
                                except:
                                    #print 'debug ' + '-'*40
                                    pass
            self.queue.task_done()
            
class ThreadUrlReadScores(threading.Thread):
    """Threaded Url Grab"""
    def __init__(self, queue):
        threading.Thread.__init__(self)
        self.queue = queue

    def run(self):
        while True:
            vhost = self.queue.get()
            #http://www.betexplorer.com/soccer/argentina/primera-division/nextmatch.php?matchid=1393301&timezone=?4
            #http://www.betexplorer.com/soccer/argentina/primera-division/matchdetails.php?matchid=1393301&timezone=?4
            #vhost = vhost.replace('nextmatch.php','matchdetails.php')
            attempts = 0
            while attempts < 10: 
                try:
                    html = urllib2.urlopen(vhost.replace('nextmatch.php','matchdetails.php'), None)
                    attempts += 1
                    break
                except urllib2.HTTPError, error: 
                    print "*** WARNING: Connection ERROR. Reconnecting after 4 sec..."
                    html = ''
                    time.sleep(4)
                    attempts += 1
                    
                print '*** ERROR: Unable to connect after ' + str(attempts) + ' tries!'
            soup = BeautifulSoup(''.join(html))
            if soup.find('th', attrs={'class':'right scorecell'}):
                rightscorecell = soup.find('th', attrs={'class':'right scorecell'}).text
                leftscorecell = soup.find('th', attrs={'class':'left scorecell'}).text
                
                if int(rightscorecell) < int(leftscorecell):
                    resscores.update({scrlinks[vhost]:'V'})
                elif int(rightscorecell) > int(leftscorecell):
                    resscores.update({scrlinks[vhost]:'H'})
                elif int(rightscorecell) == int(leftscorecell):
                    resscores.update({scrlinks[vhost]:'X'})
                
            self.queue.task_done()
            
def __request(url):
    request = url
    print 'DEBUG: __request: ' + request
    attempts = 0
    while attempts < 4:
        try:
            conn = urllib2.urlopen(request, None)
            attempts += 1
            break
        except urllib2.HTTPError, error: 
            print "*** WARNING: Connection ERROR. Reconnecting after 4 sec..."
            conn = ''
            time.sleep(4)
            attempts += 1
            
        print '*** ERROR: Unable to connect after ' + str(attempts) + ' tries!'
        
    return conn

start = time.time()

def requestURL(vURL):
    attempts = 0
    while attempts < 10:
        try:
            response = browser.load(vURL)
            attempts += 1
            break
        except:
            print '*** WARNING: Exception in address: ' + vURL
            print 'DEBUG: trying again after 4 sec...'
            response = ''
            time.sleep(4)
            attempts += 1
    return browser.html
    
def LoadListFromFile(filename):
    result = []
    fileIn = open(filename, 'r')
    for line in fileIn:
        result.append(line.strip())
    fileIn.close()
    return result

#========================================================================

parser = optparse.OptionParser()
parser.add_option('-p', '--path', dest='path', help='path to csv') # path
parser.add_option('-d', '--date', dest='date', help='needed date') # date
parser.add_option('-t', '--timezone', dest='timezone', help='needed timezone') # timezone

options, args = parser.parse_args()

if not options.timezone: # timezone can be omited
    options.timezone = '+0'
tz = '?timezone=' + str(options.timezone)
if not os.path.exists('SoccerPastRes'):
    os.mkdir('SoccerPastRes')
if not debugF:
    for f in os.listdir('./SoccerPastRes'):
        os.remove('./SoccerPastRes/' + f)
if not options.date: # date can be omited
    curDate = date.today()
else:
    
    DateTemp = map(int, str(options.date).split('/'))
    curDate = date(DateTemp[2], DateTemp[0], DateTemp[1])
    
months = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
myDate = date(2012,7,3)
print 'DEBUG: current date: ' + str(curDate)

if (curDate.year - 1)% 4 == 0 or (curDate.year - 2)% 4 ==0:
    twentymonthago = curDate - datetime.timedelta(weeks=87.0238) #86.9047
else:
    twentymonthago = curDate - datetime.timedelta(weeks=86.9047)
print 'DEBUG: twenty months ago: ' + str(twentymonthago)

#print (curDate - myDate).days
#print curDate - datetime.timedelta(days=1)
#if os.path.exists(options.path + 'res.csv'):
#    os.remove(options.path + 'res.csv')

#browser = Browster()
#browser.create_webview(show=False)
vpath = str(options.path).replace('\\', '\\\\')

file_name = vpath + 'Soccer.xlsx'
#print file_name
excel = Dispatch('Excel.Application')
excel.Visible = False  #If we want to see it change, it's fun
workbook = excel.Workbooks.Open(file_name)
workBook = excel.ActiveWorkbook
sheets = workBook.Sheets
sheet = sheets('Results')
a_row = sheet.Range("A" + str(sheet.UsedRange.rows.Count)).End(-4162).Row + 2
c_row = sheet.Range("C" + str(sheet.UsedRange.rows.Count)).End(-4162).Row + 2
#print a_row, c_row
crow = max(a_row, c_row, 5)
#print crow
sheet.Activate()
sheet.Cells(crow,1).Value = str(curDate.month) + '/' + str(curDate.day) + '/' + str(curDate.year)
sheet.Cells(crow+2,1).Value = weekDay()
sheet.Cells(crow+4,1).Value = curDate.isocalendar()[1]

if not debugF:
    countryList = LoadListFromFile('SoccerCountryList.txt')
    #storedhtml = requestURL('http://www.betexplorer.com/soccer/')
    #?timezone=+3
    storedhtml = __request('http://www.betexplorer.com/soccer/' + tz)
    soup = BeautifulSoup(''.join(storedhtml)) #parse html source

    # we will monitor this text. Once it's not present on the page - means ajax is done
    countries = soup.find('div', attrs={'class':'countries'})
    lookupPoints = countries.findAll('a')
    print 'DEBUG: Found ' + str(len(lookupPoints)) + ' countries'
    print 'DEBUG: Site should be analysed only for: ' 
    print countryList
    countryLinks = []
    for lookupPoint in lookupPoints:
        if lookupPoint['href'][lookupPoint['href'][:-1].rfind('/')+1:-1] in countryList:
            countryLinks.append('http://www.betexplorer.com' + str(lookupPoint['href'] + tz))
    
    firstLinks = []    
    for i in range(len(countryLinks)):
    #for i in range(0,3):
        t = ThreadUrlRead(queue)
        t.setDaemon(True)
        t.start()
    for countryLink in countryLinks:
        queue.put(countryLink)
    queue.join()
    
    print 'DEBUG: First links obtained (' + str(len(firstLinks)) + ')'
    #print firstLinks

    queue1 = Queue.Queue()
    GameLinks = []
    for i in range(len(firstLinks)):
    #for i in range(0,1):
        t = ThreadUrlReadGames(queue1)
        t.setDaemon(True)
        t.start()
    for firstLink in firstLinks:
        queue1.put(firstLink)
    queue1.join()
    
    print 'Game links obtained (' + str(len(GameLinks)) + '): '
    print GameLinks
    
    queue2 = Queue.Queue()
    #for i in range(len(GameLinks)):
    for i in range(0,11):
        t = ThreadUrl(queue2)
        t.setDaemon(True)
        t.start()
    for GameLink in GameLinks:
        queue2.put(GameLink)
    queue2.join()

    print 'DEBUG: Results saved'
    
# now we can get html source
i = 0
for f in os.listdir('./SoccerPastRes'):
    i += 1
    #if i == 2:
    #    break
    #print f
    html = open('./SoccerPastRes/' + f,'r')
    soup = BeautifulSoup(''.join(html)) #parse html source

    #[9/11/2012 8:19:07 PM] Dragan: This time we are doing soccer, and the website is betexplorer.com.
    
    #[9/11/2012 8:20:19 PM] Dragan: Col. C and D are the Home team and the Visiting team. 
    #Note that its reversed from MLB and NHL. There it was Vis and Home teams.
    #players = soup.find('h1', attrs={'class':'nextmatch'}).find('span', attrs={'class':'fleft'}).find(text=True)
    players = soup.find('h1', attrs={'class':'nextmatch'}).find('span', attrs={'class':'fleft'}).text
    try:
        print 'DEBUG: current game: ' + str(players.encode('utf-8'))
    except:
        pass
    vC = players.split(' - ')[0].strip()
    vD = players.split(' - ')[1].strip()
    
    
    #Next thing is columns E and F. E says FAV (favorite), and the returning value should be "H" or "V" home or visitor. 
    #in this example Batna - USM home team has a lower coefficient so home team is the favorite, so 
    #column E should be "H".If the visitor has the lower coefficient, then "V".
    #Note: in some rare cases you will see the same coefficient like in game 3, that game sould be counted but leave the col E blank. the tool will 
    #probably get confused if it runs into 
    #such and example without you programming it what to do. Right?

    #ok. column looks for the team that is NOT the favorite.
    #OU3 means is the coefficient Over or Under 3.00 If it's Over then "O", if Under then "U".
    #Let me show you an example.
    #Batna - USM col. E is H, cOL f IS "u", 
    #rIED - Admira col. E is H, Col F is "O" cuz it above 3.00
    #what about exact 3.00?
    #That is "O".
    #what about if blank for E?
    #Then it's "U". If both coefficients are the same, then they will ALWAYS be below 3.00, usually at about 2.5 or 2.6. so, it's "U".

    favorTag = soup.find('ul', attrs={'class':'add-to'})
    if not favorTag:
        favorTag = soup.find('ul', attrs={'class':'add-to live'})
    favorlis = favorTag.findAll('li')
    homefavor = float(favorlis[1].text)
    visfavor = float(favorlis[3].text)
    if homefavor < visfavor:
        vE = 'H'
        if visfavor < 3.00:
            vF = 'U'
        elif visfavor >= 3.00:
            vF = 'O'
    elif homefavor > visfavor:
        vE = 'V'
        if homefavor < 3.00:
            vF = 'U'
        elif homefavor >= 3.00:
            vF = 'O'
    else:
        vE = ''
        vF = 'U'

    #[9/11/2012 8:23:59 PM] Dragan: Col. G is the country. In the template I've sent you the game is from Ireland. 
    #You can use the abbreviations "Ire", or you can copy the whole name from the website - "Ireland", 
    #whichever is easier on you. 
    #There will be a lot of countries, for example: "Rus" or "Russia", "Bra" or "Brasil". 
    #Both options are fine as long as I can understand which country it is.
    mlinkTag = soup.find('ul', attrs={'id':'location'})
    mLinks = mlinkTag.findAll('a')
    for mLink in mLinks:
        mCountry = mLink['href']
    vG = mCountry.split('/')[2]
    
    #result-table team-matches
    WDLTables = soup.findAll('table', attrs={'class':'result-table team-matches'})
    
    #[9/11/2012 8:28:21 PM] Dragan: Col. I-N are the same as col. P-S in MLB, except here we also have the draw column. 
    #So, I-home win, J-home draw, K-home loss, L-visitor win, M-visitor draw, N-visitor loss. 
    #In the template home team won it's last 2 games, and the visiting team lost its last 2 games.
    vI = 0
    vJ = 0
    vK = 0
    IJKst = False
    vbreak = False
    for WDLTable in WDLTables[:1]:
        rows = WDLTable.findAll('tr')            
        for tr in rows: 
            if 'No matches found' in tr.find(text=True):
                print 'WARNING: Seems like last-results table in blank...'
                vI = -1
                vJ = -1
                vK = -1
            else:
                cols = tr.findAll('td')
                for td in cols[:1]: 
                    if IJKst:
                        if IJKprev <> td:
                            vbreak = True
                            break
                    IJKprev = td
                    if td.find('span', attrs={'class':'form-bg form-w'}):
                        vI += 1
                    elif td.find('span', attrs={'class':'form-bg form-d'}):
                        vJ += 1
                    elif td.find('span', attrs={'class':'form-bg form-l'}):
                        vK += 1
                    IJKst = True
            if vbreak:
                break

    vL = 0
    vM = 0
    vN = 0
    LMNst = False
    vbreak = False
    for WDLTable in WDLTables[1:]:
        rows = WDLTable.findAll('tr')            
        for tr in rows: 
            if 'No matches found' in tr.find(text=True):
                print 'WARNING: Seems like last-results table in blank...'
                vL = -1
                vM = -1
                vN = -1
            else:
                cols = tr.findAll('td')
                for td in cols[:1]: 
                    if LMNst:
                        if LMNprev <> td:
                            vbreak = True
                            break
                    LMNprev = td
                    if td.find('span', attrs={'class':'form-bg form-w'}):
                        vL += 1
                    elif td.find('span', attrs={'class':'form-bg form-d'}):
                        vM += 1
                    elif td.find('span', attrs={'class':'form-bg form-l'}):
                        vN += 1
                    LMNst = True
            if vbreak:
                break

    #+[9/11/2012 8:38:57 PM] Dragan: Col. P-U are the same as col. I-N exept that here I am only interested 
    #in the games where today's home team was also the home team, 
    #and today's visiting team was also the visiting team. 
    #In other words, all the games where today's home team was the visitor should be excluded, 
    #and all the games where today's visiting team was the home team should be excluded. 
    #In this example col. P is "1", which means that today's home team won its last home game, 
    #col. U is "1", which means that today's visitng team lost its last game as a visitor.
   
    vP = 0
    vQ = 0
    vR = 0
    PQRst = False
    vbreak = False
    for WDLTable in WDLTables[:1]:
        rows = WDLTable.findAll('tr')            
        for tr in rows: 
            if 'No matches found' in tr.find(text=True):
                print 'WARNING: Seems like last-results table in blank...'
                vP = -1
                vQ = -1
                vR = -1
            else:
                cols = tr.findAll('td') 
                for td in cols[:1]: 
                    if (len(vC) <= 16 and cols[1].find(text=True).strip() == vC) or (len(vC) > 16 and cols[1].find(text=True)[:16].strip() == vC[:16]):
                        if PQRst:
                            if PQRprev <> td:
                                vbreak = True
                                break
                        PQRprev = td
                        if td.find('span', attrs={'class':'form-bg form-w'}):
                            vP += 1
                        elif td.find('span', attrs={'class':'form-bg form-d'}):
                            vQ += 1
                        elif td.find('span', attrs={'class':'form-bg form-l'}):
                            vR += 1
                        PQRst = True
            if vbreak:
                break

    vS = 0
    vT = 0
    vU = 0
    STUst = False
    vbreak = False
    for WDLTable in WDLTables[1:]:
        rows = WDLTable.findAll('tr')            
        for tr in rows: 
            if 'No matches found' in tr.find(text=True):
                print 'WARNING: Seems like last-results table in blank...'
                vS = -1
                vT = -1
                vU = -1
            else:
                cols = tr.findAll('td')
                for td in cols[:1]: 
                    if (len(vD) <= 16 and cols[2].find(text=True).strip() == vD) or (len(vD) > 16 and cols[2].find(text=True)[:16].strip() == vD[:16]):
                        if STUst:
                            if STUprev <> td:
                                vbreak = True
                                break
                        STUprev = td
                        if td.find('span', attrs={'class':'form-bg form-w'}):
                            vS += 1
                        elif td.find('span', attrs={'class':'form-bg form-d'}):
                            vT += 1
                        elif td.find('span', attrs={'class':'form-bg form-l'}):
                            vU += 1
                        STUst = True
            if vbreak:
                break

    
    #[9/11/2012 8:41:07 PM] Dragan: Col. W-Z is looking only into draws, nothing else. 
    #It tells me when was the last time a team played a draw.
    #[9/11/2012 8:42:14 PM] Dragan: Forgot to mention, the tool should look only into the last 10 games. 
    #This goes for all columns except for AL-AO.
    
    #[9/11/2012 9:00:02 PM] Dragan: Col. W - How many games has it been since the home team had a draw? 
    #If there hasn't been a draw in the last 10 games, then the return value is "11".
    vW = 0
    vbreak = False
    wasdraw = False
    for WDLTable in WDLTables[:1]:
        rows = WDLTable.findAll('tr')            
        for tr in rows: 
            if 'No matches found' in tr.find(text=True):
                print 'WARNING: Seems like last-results table in blank...'
                wasdraw = True
                vW = -1
            else:
                cols = tr.findAll('td') 
                for td in cols[:1]: 
                    vW +=1
                    if td.find('span', attrs={'class':'form-bg form-d'}):
                        wasdraw = True
                        vbreak = True
                        break
                    
            if vbreak:
                break
    if vW == 0 or not wasdraw:
        vW = 11
    
    #[9/11/2012 9:01:03 PM] Dragan: Col. X - How many games has it been since the visiting team had a draw? 
    #In the template, it say that it has been 4 games since the visiting team had a draw.
    vX = 0
    vbreak = False
    wasdraw = False
    for WDLTable in WDLTables[1:]:
        rows = WDLTable.findAll('tr')            
        for tr in rows: 
            if 'No matches found' in tr.find(text=True):
                print 'WARNING: Seems like last-results table in blank...'
                wasdraw = True
                vX = -1
            else:
                cols = tr.findAll('td') 
                for td in cols[:1]: 
                    vX +=1
                    if td.find('span', attrs={'class':'form-bg form-d'}):
                        wasdraw = True
                        vbreak = True
                        break
                    
            if vbreak:
                break
    if vX == 0 or not wasdraw:
        vX = 11

    #[9/11/2012 9:02:45 PM] Dragan: Col. Y - How many games has it been since the home team had a draw as a home team? 
    #So, exclude the games where the home team was the visitor. 
    #In the template there has not been a draw in the last 10 games, so the value is "11".
    vY = 0
    vbreak = False
    wasdraw = False
    for WDLTable in WDLTables[:1]:
        rows = WDLTable.findAll('tr')            
        for tr in rows: 
            if 'No matches found' in tr.find(text=True):
                print 'WARNING: Seems like last-results table in blank...'
                wasdraw = True
                vY = -1
            else:
                cols = tr.findAll('td') 
                for td in cols[:1]: 
                    if (len(vC) <= 16 and cols[1].find(text=True).strip() == vC) or (len(vC) > 16 and cols[1].find(text=True)[:16].strip() == vC[:16]):
                        vY +=1
                        if td.find('span', attrs={'class':'form-bg form-d'}):
                            wasdraw = True
                            vbreak = True
                            break
                    
            if vbreak:
                break
    if vY == 0 or not wasdraw:
        vY = 11
    
    #[9/11/2012 9:04:21 PM] Dragan: Col. Z - How many games has it been since the visiting team had a draw as a visitor? 
    #So exclude all the games where the visiting team was the home team. 
    #In our template, 3 games ago the visiting team (as a visitor) had a draw.
    vZ = 0
    vbreak = False
    wasdraw = False
    for WDLTable in WDLTables[1:]:
        rows = WDLTable.findAll('tr')            
        for tr in rows: 
            if 'No matches found' in tr.find(text=True):
                print 'WARNING: Seems like last-results table in blank...'
                wasdraw = True
                vZ = -1
            else:
                cols = tr.findAll('td') 
                for td in cols[:1]: 
                    if (len(vD) <= 16 and cols[2].find(text=True).strip() == vD) or (len(vD) > 16 and cols[2].find(text=True)[:16].strip() == vD[:16]):
                        vZ +=1
                        if td.find('span', attrs={'class':'form-bg form-d'}):
                            wasdraw = True
                            vbreak = True
                            break
            if vbreak:
                break
    if vZ == 0 or not wasdraw:
        vZ = 11
    
    
    #[9/11/2012 9:14:46 PM] Dragan: AL - AO should look into the past 2 seasons, the current and the previos one. 
    #2012-13, and 2011-12.
    HHTable = soup.find('table', attrs={'class':'result-table'})
    
    if HHTable.find('td', attrs={'class':'nomatch first-cell last-cell nobr'}):
        print '*** WARNING: No head-to-head data found. cols AG-AI skipped...'
        vAL = '-'
        vAM = '-'
        vAN = '-'
    else:
        #[9/11/2012 9:18:30 PM] Dragan: AL - How did the last game between today's 2 teams end?
        #( H-Home win, X-Draw, V-Visitor won). 
        #AL looks only into the last game where today's home teams was also home, 
        #and today's visitor was also visitor.

        rows = HHTable.findAll('tr') 
        vAL = ''
        vbreak = False
        tdcounter = 0
        for tr in rows: 
            cols = tr.findAll('td') 
            for td in cols[:1]: 
                
                hhdatearr = cols[7].find(text=True).split('.')
                hhdate = date(int(hhdatearr[2]), int(hhdatearr[1]), int(hhdatearr[0]))
                legpass = False
                for firstlegue in firstlegues:
                    if firstlegue[:-1] in cols[6].find('a')['href']:
                        legpass = True
                if legpass:
                    
                    if hhdate >= twentymonthago:
                        tdcounter+=1
                        if (len(vC) <= 16 and td.find(text=True).strip() == vC) or (len(vC) > 16 and td.find(text=True)[:16].strip() == vC[:16]):
                            vPoints = cols[2].find(text=True).split(':')
                            if vPoints[0] > vPoints[1]:
                                vAL = 'H'
                                vbreak = True
                                break
                            elif vPoints[0] < vPoints[1]:
                                vAL = 'V'
                                vbreak = True
                                break
                            elif vPoints[0] == vPoints[1]:
                                vAL = 'X'
                                vbreak = True
                                break
            if vbreak:
                break
        if tdcounter == 1:
            thesame = True
        else:
            thesame = False
            
        #[9/11/2012 9:20:01 PM] Dragan: AN - Looks only into the last game between these 2 teams regardless 
        #of who was the visitor and who was the home team. 
        #In our template, the visitor won the last game these 2 teams met.
        vAN = ''
        vAM = ''
        vbreak = False
        for tr in rows: 
            cols = tr.findAll('td') 
            for td in cols[:1]: 
                hhdatearr = cols[7].find(text=True).split('.')
                hhdate = date(int(hhdatearr[2]), int(hhdatearr[1]), int(hhdatearr[0]))
                legpass = False
                for firstlegue in firstlegues:
                    if firstlegue[:-1] in cols[6].find('a')['href']:
                        legpass = True
                if legpass:
                    if hhdate >= twentymonthago:
                        vPoints = cols[2].find(text=True).split(':')
                        if vPoints[0] > vPoints[1]:
                            vAN = 'H'
                            vAM = 'H'
                            vbreak = True
                            break
                        elif vPoints[0] < vPoints[1]:
                            vAN = 'V'
                            vAM = 'V'
                            vbreak = True
                            break
                        elif vPoints[0] == vPoints[1]:
                            vAN = 'X'
                            vAM = 'X'
                            vbreak = True
                            break
            if vbreak:
                break
        if thesame:
            vAN = ''
        else:
            vAM = ''
        
        
    vAB = 0
    vbreak = False
    wasdraw = False
    for WDLTable in WDLTables[:1]:
        rows = WDLTable.findAll('tr')            
        for tr in rows: 
            cols = tr.findAll('td') 
            for td in cols[:1]: 
                if 'No matches found' in tr.find(text=True):
                    print 'WARNING: Seems like last-results table is blank...'
                    wasdraw = True
                    vAB = -1
                else:
                    vPoints = cols[3].find(text=True).split(':')
                    coef1 = float(cols[4].find(text=True))
                    coef2 = float(cols[6].find(text=True))
                    if coef1 >=2.0 and coef2 >= 2.0:
                        vAB+=1
                        if int(vPoints[0]) == int(vPoints[1]):
                            wasdraw = True
                            vbreak = True
                            break
            if vbreak:
                break
        if vbreak:
            break
    if not wasdraw:
        vAB = 0
        
    vAC = 0
    vbreak = False
    wasdraw = False
    for WDLTable in WDLTables[:1]:
        rows = WDLTable.findAll('tr')            
        for tr in rows: 
            if 'No matches found' in tr.find(text=True):
                print 'WARNING: Seems like last-results table in blank...'
                wasdraw = True
                vAC = -1
            else:
                cols = tr.findAll('td') 
                for td in cols[:1]: 
                    if (len(vC) <= 16 and cols[1].find(text=True).strip() == vC) or (len(vC) > 16 and cols[1].find(text=True)[:16].strip() == vC[:16]):
                        vPoints = cols[3].find(text=True).split(':')
                        coef1 = float(cols[4].find(text=True))
                        coef2 = float(cols[6].find(text=True))
                        if coef1 >=2.0 and coef2 >= 2.0:
                            vAC+=1
                            if int(vPoints[0]) == int(vPoints[1]):
                                wasdraw = True
                                vbreak = True
                                break
            if vbreak:
                break
        if vbreak:
            break
    if not wasdraw:
        vAC = 0

    vAD = 0
    vbreak = False
    wasdraw = False
    for WDLTable in WDLTables[1:]:
        rows = WDLTable.findAll('tr')            
        for tr in rows: 
            if 'No matches found' in tr.find(text=True):
                print 'WARNING: Seems like last-results table in blank...'
                wasdraw = True
                vAD = -1
            else:
                cols = tr.findAll('td') 
                for td in cols[1:2]: 
                    vPoints = cols[3].find(text=True).split(':')
                    coef1 = float(cols[4].find(text=True))
                    coef2 = float(cols[6].find(text=True))
                    if coef1 >=2.0 and coef2 >= 2.0:
                        vAD+=1
                        if int(vPoints[0]) == int(vPoints[1]):
                            wasdraw = True
                            vbreak = True
                            break
            if vbreak:
                break
        if vbreak:
            break
    if not wasdraw:
        vAD = 0
        
    vAE = 0
    vbreak = False
    wasdraw = False
    for WDLTable in WDLTables[1:]:
        rows = WDLTable.findAll('tr')            
        for tr in rows: 
            if 'No matches found' in tr.find(text=True):
                print 'WARNING: Seems like last-results table in blank...'
                wasdraw = True
                vAE = -1
            else:
                cols = tr.findAll('td') 
                for td in cols[1:2]: 
                    if (len(vD) <= 16 and cols[2].find(text=True).strip() == vD) or (len(vD) > 16 and cols[2].find(text=True)[:16].strip() == vD[:16]):
                        vPoints = cols[3].find(text=True).split(':')
                        coef1 = float(cols[4].find(text=True))
                        coef2 = float(cols[6].find(text=True))
                        if coef1 >=2.0 and coef2 >= 2.0:
                            vAE+=1
                            if int(vPoints[0]) == int(vPoints[1]):
                                wasdraw = True
                                vbreak = True
                                break
            if vbreak:
                break
        if vbreak:
            break
    if not wasdraw:
        vAE = 0
        
    try:
        sheet.Range('C'+str(crow)).value = vC
    except:
        try:
            sheet.Range('C'+str(crow)).value = vC.encode('utf-8')
        except:
            sheet.Range('C'+str(crow)).value = str(vC)
        
    try:
        sheet.Range('D'+str(crow)).value = vD
    except:
        try:
            sheet.Range('D'+str(crow)).value = vD.encode('utf-8')
        except:
            sheet.Range('D'+str(crow)).value = str(vD)
        
    sheet.Range('G'+str(crow)).value = vG
    sheet.Range('E'+str(crow)).value = vE
    sheet.Range('F'+str(crow)).value = vF

    if str(vI) == str('0'):
        sheet.Range('I'+str(crow)).value = ''
    else:
        sheet.Range('I'+str(crow)).value = str(vI)
    if str(vJ) == str('0'):
        sheet.Range('J'+str(crow)).value = ''
    else:
        sheet.Range('J'+str(crow)).value = str(vJ)
    if str(vK) == str('0'):
        sheet.Range('K'+str(crow)).value = ''
    else:
        sheet.Range('K'+str(crow)).value = str(vK)
    if str(vL) == str('0'):
        sheet.Range('L'+str(crow)).value = ''
    else:
        sheet.Range('L'+str(crow)).value = str(vL)
    if str(vM) == str('0'):
        sheet.Range('M'+str(crow)).value = ''
    else:
        sheet.Range('M'+str(crow)).value = str(vM)
    if str(vN) == str('0'):
        sheet.Range('N'+str(crow)).value = ''
    else:
        sheet.Range('N'+str(crow)).value = str(vN)

    if str(vP) == str('0'):
        sheet.Range('P'+str(crow)).value = ''
    else:
        sheet.Range('P'+str(crow)).value = str(vP)
    if str(vQ) == str('0'):
        sheet.Range('Q'+str(crow)).value = ''
    else:
        sheet.Range('Q'+str(crow)).value = str(vQ)
    if str(vR) == str('0'):
        sheet.Range('R'+str(crow)).value = ''
    else:
        sheet.Range('R'+str(crow)).value = str(vR)
    if str(vS) == str('0'):
        sheet.Range('S'+str(crow)).value = ''
    else:
        sheet.Range('S'+str(crow)).value = str(vS)
    if str(vT) == str('0'):
        sheet.Range('T'+str(crow)).value = ''
    else:
        sheet.Range('T'+str(crow)).value = str(vT)
    if str(vU) == str('0'):
        sheet.Range('U'+str(crow)).value = ''
    else:
        sheet.Range('U'+str(crow)).value = str(vU)

    sheet.Range('W'+str(crow)).value = str(vW)
    sheet.Range('X'+str(crow)).value = str(vX)
    sheet.Range('Y'+str(crow)).value = str(vY)
    sheet.Range('Z'+str(crow)).value = str(vZ)

    if str(vAB) == str('0'):
        sheet.Range('AB'+str(crow)).value = '11'
    else:
        sheet.Range('AB'+str(crow)).value = str(vAB)
    if str(vAC) == str('0'):
        sheet.Range('AC'+str(crow)).value = '11'
    else:
        sheet.Range('AC'+str(crow)).value = str(vAC)
    if str(vAD) == str('0'):
        sheet.Range('AD'+str(crow)).value = '11'
    else:
        sheet.Range('AD'+str(crow)).value = str(vAD)
    if str(vAE) == str('0'):
        sheet.Range('AE'+str(crow)).value = '11'
    else:
        sheet.Range('AE'+str(crow)).value = str(vAE)

    sheet.Range('AG'+str(crow)).value = str(vAL)
    sheet.Range('AH'+str(crow)).value = str(vAM)
    sheet.Range('AI'+str(crow)).value = str(vAN)
    
    sheet.Range('AP'+str(crow)).value = curDate
    sheet.Range('AQ'+str(crow)).value = f[:-4].replace('^',':').replace('_','/').replace('+','?')
    
    crow +=1


scrlinks = {}
for chkrow in range(5,crow):
    if sheet.Range('AL'+str(chkrow)).value is None:
        if sheet.Range('AQ'+str(chkrow)).value is not None:
            try:
                rowdatetext = map(int, str(sheet.Range('AP'+str(chkrow)).text).split('/'))
                rowdate = date(rowdatetext[2], rowdatetext[0], rowdatetext[1])
                if rowdate < date.today():
                    scrlinks.update({sheet.Range('AQ'+str(chkrow)).value:chkrow})
                    
            except:
                pass
print 'DEBUG: Trying to fill past scores: '
print scrlinks.keys()
resscores ={}
queue3 = Queue.Queue()
for i in range(len(scrlinks.keys())):
    t = ThreadUrlReadScores(queue3)
    t.setDaemon(True)
    t.start()
for scrlink in scrlinks.keys():
    queue3.put(scrlink)
queue3.join()

if len(resscores) > 0:
    vscoresrows = resscores.viewkeys()
    vscoresvalues = resscores.viewvalues()
    for vscoresrow in vscoresrows:
        sheet.Range('AL'+str(vscoresrow)).value = resscores[vscoresrow]
else:
    print '*** WARNING: no data to fill past scores'
workBook.Save
excel.Visible = True

#columns AG-AJ should be deleted, I did not want to do that cuz it would probably mess up the tool. !!!!
