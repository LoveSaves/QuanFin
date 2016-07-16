# ============================================================================================================================
# python script that uses the selenium package to automatically brows http://www.nasdaqomxnordic.com/aktier/historiskakurser
# in order to download and store historical data on the Stockholm stock-exchange.
# ============================================================================================================================
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
import os
import subprocess
import shutil
import distutils.core
import datetime
import time

OSX = False  # Set to true if you are working on your mac at home.
# mydriver = webdriver.Firefox()
mydriver = webdriver.Chrome()
# mydriver = webdriver.Remote( browser_name="chrome") #, platform="any")

syear = '2000'
smonth = '09'
sday = '20'
startdate = syear + '-' + smonth + '-' + sday

stokname = []
xpaths = []

outfile = open("D:\QuanFin\Data\\listofstocks.csv", 'w')
# ========================================================================
# The list of large companies on the Stockholm stock-market:
# ========================================================================
stokname.append(("CSE39406","Sparinvest Index Globale AKT MIN RIS KL")) #Sparinvest Index Globale AKT MIN RIS KL
stokname.append(("CSE90733","Nordea Invest Globale Aktier Indeks")) #Nordea Inv. Globale Aktier Indeks
#stokname.append("A.P. Møller - Mærsk B")
#stokname.append("Carlsberg B")
#stokname.append("Chr. Hansen Holding")
#stokname.append("Coloplast B")
#stokname.append("Danske Bank")
#stokname.append("DK0060079531") #DSV
#stokname.append("FLSmidth & Co.")
#stokname.append("Genmab")
#stokname.append("GN Store Nord")
#stokname.append("ISS")
#stokname.append("DK0010307958") #Jyske Bank
#stokname.append("Lundbeck")
#stokname.append("NDA DKK") #Nordea Bank
#stokname.append("Novo Nordisk B")
#stokname.append("NZYM B")
#stokname.append("Pandora")
#stokname.append("TDC")
#stokname.append("Tryg")
#stokname.append("Vestas Wind Systems")
#stokname.append("William Demant Holding")



# Here we get the date of today
eyear = str(datetime.datetime.now().year)
eday = str(datetime.datetime.now().day - 1)
emonth = str(datetime.datetime.now().month)

if datetime.datetime.now().day - 1 < 10: eday = '0' + eday
if datetime.datetime.now().month < 10: emonth = '0' + emonth

enddate = eyear + '-' + emonth + '-' + eday

xpaths.append('//*[@id="instSearchHistorical"]')
xpaths.append('//*[@id="FromDate"]')
xpaths.append('//*[@id="ToDate"]')

xpaths.append('//*[@id="exportExcel"]')

i = 0

# ================================================================================================================
# Loop over all big and medium sized companies at the Stockholm stock-market. Downloading historical data of the
# stock from a date specified by the user to today.
# ================================================================================================================
for name in stokname:

    mydriver.get('http://www.nasdaqomxnordic.com/Funds/Historical_Prices?Instrument='+name[0])
    time.sleep(2)

    #mydriver.find_element_by_xpath('//*[@id="instSearchHistorical"]').send_keys(name)
    #time.sleep(1)

    #mydriver.find_element_by_xpath('//*[@id="instSearchHistorical"]').send_keys(Keys.ARROW_DOWN)

    #time.sleep(1)
    #mydriver.find_element_by_xpath('//*[@id="instSearchHistorical"]').send_keys(Keys.ENTER)
    #time.sleep(1)
    mydriver.find_element_by_xpath('//*[@id="FromDate"]').clear()
    time.sleep(1)
    mydriver.find_element_by_xpath('//*[@id="FromDate"]').send_keys(startdate)
    time.sleep(1)
    mydriver.find_element_by_xpath('//*[@id="ToDate"]').clear()
    time.sleep(1)
    mydriver.find_element_by_xpath('//*[@id="ToDate"]').send_keys(enddate)
    mydriver.find_element_by_xpath('//*[@id="ToDate"]').send_keys(Keys.TAB)
    time.sleep(2)

    mydriver.find_element_by_xpath('//*[@id="exportExcel"]').click()

    index = str(i + 1)
    savefile = 'D:\\QuanFin\\Data\\' + name[1]
    savefile = savefile + '.csv'

    k = 0
    time.sleep(2)

    downloadfile = ""
    for line in os.listdir("C://Users//kim.simonsen//Downloads//."):
        if line.endswith(".csv") and k == 0: downloadfile = line
        k = k + 1

    downloadfile = downloadfile.strip('\n')
    print(downloadfile)
    time.sleep(1.0)

    downl = "C:\\Users\\kim.simonsen\\Downloads\\" + downloadfile
    shutil.move(downl, savefile)

    i = i + 1
    print(name)

    outfile.write(name[0] + "," + name[1] + "\n")
# Here we close the driver after all the data has been downloaded
mydriver.close()
outfile.close()
