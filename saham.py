import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import mplcursors as mplc
import datetime as dt
import pandas as pd
import numpy as np
import timedelta
import requests
import shutil
import math
import time
import sys
import os
from mplfinance.original_flavor import candlestick_ohlc
from matplotlib import ticker
from selenium import webdriver
from pandas.tseries.offsets import BDay

# pip install xlrd pandas, requests, matplotlib, timedelta, selenium, mplcursors
# pip install https://github.com/matplotlib/mpl_finance/archive/master.zip

# Global Variables
flag = 0
chromeInitialization = False

#######################################################
# Date Function                                       #
# stringType = 0, format example = 20200916           #
# stringType = 1, format example = September 16, 2020 #
# stringType = 2, format example = 16/09/2020         #
# stringType = 3, format example = 2020-09-16         #
#######################################################
def getDate(bDaysBefore, stringType):
    global flag
    today = dt.datetime.today()
    if(stringType == 0):
        date = (today - BDay(bDaysBefore)).strftime('%Y%m%d')
        return date
    elif(stringType == 1):
        date = (today - BDay(bDaysBefore)).strftime('%B %#d, %Y')
        currMonth = today.strftime('%B')
        dateMonth = (today - BDay(bDaysBefore)).strftime('%B')
        if(currMonth != dateMonth and flag == 0):
            flag = 1
        return date
    elif(stringType == 2):
        date = (today - BDay(bDaysBefore)).strftime('%d/%m/%Y')
        return date
    elif(stringType == 3):
        date = (today - BDay(bDaysBefore)).strftime('%Y-%m-%d')
        return date

# Download the most recent stock market data from IDX using google chrome
def getData(date):
    global googleChrome
    global chromeInitialization
    global flag
    if (chromeInitialization == False):
        chromeInitialization = True
        url = "https://www.idx.co.id/data-pasar/ringkasan-perdagangan/ringkasan-saham"
        googleChrome = webdriver.Chrome()
        # googleChrome.minimize_window()
        googleChrome.get(url)
    googleChrome.find_element_by_id("dateFilter").click()
    string = "//span[@aria-label = '" + date + "']"
    if(flag == 1):
        googleChrome.find_element_by_xpath("//span[@class = 'flatpickr-prev-month']").click()
        flag = 2
    print(string)
    time.sleep(0.5)
    googleChrome.find_element_by_xpath(string).click()
    print("masuk")
    googleChrome.find_element_by_xpath("/html/body/main/div[2]/div/div[3]/a/span").click()

#########################################################################
# Copy the file from \downloads folder to \python folder, if necessary  #
# If th downloaded file is < 700KB, then get another valid data from    #
# 1 previous business day. This will make sure to have the 10 most      #
# reason stock price file downlaoded                                    #
#########################################################################
def checkData():
    count = 0
    max = 25
    while(count < max):
        date = getDate(count,0)
        location = r'C:\Users\fenti\Downloads\Ringkasan Saham-'
        destination = r'C:\Users\fenti\Documents\Python\Ringkasan Saham-'
        fileNameOriginal = location + date + '.xlsx'
        fileNameTarget = destination + date + '.xlsx'
        date = getDate(count,1)
        if(os.path.isfile(fileNameOriginal) == False and os.path.isfile(fileNameTarget) == False):
            getData(date)
            while(os.path.isfile(fileNameOriginal) == False):
                time.sleep(1)
            if(os.stat(fileNameOriginal).st_size < 700000):
                count += 1
                max += 1
            else:
                shutil.move(fileNameOriginal, fileNameTarget)
        elif(os.path.isfile(fileNameOriginal) == True and os.path.isfile(fileNameTarget == False)):
            shutil.move(fileNameOriginal, fileNameTarget)
        if(os.path.isfile(fileNameOriginal) == True):
            max += 1
        count += 1
    if chromeInitialization == 1:
        googleChrome.close()

# Read the last 10 files and display it to the user
def displayData(stockName):
    day = 0
    openList = []
    closeList = []
    highList = []
    lowList = []
    foreignNetList = []
    deltaList = []
    dateList = []
    companyName = ""
    string = "Ringkasan Saham-" + getDate(day,0) + ".xlsx"
    for i in range (0,25):
        while(os.path.isfile(string) == False):
            day += 1
            string = "Ringkasan Saham-" + getDate(day,0) + ".xlsx"
        stockFile = pd.read_excel(string, skiprows = 1)
        dfStock = stockFile[(stockFile["Kode Saham"] == stockName)]
        if not companyName:
            companyName = (dfStock['Nama Perusahaan']).iloc[0]
        open = (dfStock['Sebelumnya']).iloc[0]
        close = (dfStock['Penutupan']).iloc[0]
        high = (dfStock['Tertinggi']).iloc[0]
        low = (dfStock['Terendah']).iloc[0]
        foreignBuy = (dfStock['Foreign Buy']).iloc[0]
        foreignSell = (dfStock['Foreign Sell']).iloc[0]
        openList.insert(0,open)
        closeList.insert(0,close)
        highList.insert(0,high)
        lowList.insert(0,low)
        foreignNetList.insert(0,foreignBuy - foreignSell)
        deltaList.insert(0,(close - open)/open)
        dateList.insert(0,getDate(day,3))
        day += 1
        string = "Ringkasan Saham-" + getDate(day,0) + ".xlsx"

    # Declaring 2 charts with 21x9 size
    fig, (axs1, axs2) = plt.subplots(2, figsize = (21,9))
    string = companyName + " (" + stockName + ")"
    fig.suptitle(string, fontweight = "bold")

    # Plot for 1st chart (top)
    stockData = [dateList, openList, highList, lowList, closeList]
    dfStockData = pd.DataFrame(stockData, index = ["Date", "Open", "High", "Low", "Close"])
    dfStockData = dfStockData.T
    dfStockData = dfStockData.set_index("Date")
    dfStockData.reset_index(inplace = True)
    xticks = []
    xticklabels = []
    for x in range(0,25,3):
        xticks.append(x)
        xticklabels.append(dateList[x])
    numOfData = len(dfStockData["Date"])
    dfStockData["Date"] = np.arange(start = 0, stop = numOfData, step = 1, dtype='int')
    candlestick_ohlc(axs1, dfStockData.values, width = 0.5, colorup = "g", colordown = "r")
    axs1.set_xticks(xticks)                 # set number of ticks
    axs1.set_xticklabels(xticklabels)       # set ticks label/name
    axs1.set_ylabel("Stock Price", fontweight = "bold")
    axs1.grid()

    # # Plot for 2nd chart (bottom)
    axs2.bar(dateList, foreignNetList, color = "orange")
    axs2.xaxis.set_major_locator(ticker.MaxNLocator(len(xticks)+1))        # set number of ticks

    axs2.set_ylabel("Foreign Net",  fontweight = "bold")
    axs2.grid()
    r = correlation(foreignNetList, deltaList, len(foreignNetList))
    power = ""
    if(abs(r) >= 0.5):
        power = "Strong Foreign Power"
    elif(abs(r) >= 0.3 and abs(r) < 0.5 ):
        power = "Medium Foreign Power"
    else:
        power = "Low Foreign Power"

    axs2.set_title("Correlation: " + str("{:.4f}".format(r)) + " (" + power + ")", loc = 'right', fontweight = "bold")

    # # axs to get correlation scatter graph
    # axs3.scatter(deltaList, foreignNetList)
    # axs3.grid()

    mplc.cursor(axs2).connect(
        "add", lambda sel: sel.annotation.set_text(foreignNetList[sel.target.index]))

    plt.show(block = False)

def correlation(deltaList, foreign, sizeOfData):
    x = 0
    y = 0
    xx = 0
    yy = 0
    xy = 0
    for i in deltaList:
        x += i
        xx += i*i
    for i in foreign:
        y += i
        yy += i*i
    for i in range(0,sizeOfData):
        xy += deltaList[i] * foreign[i]
    correlation = (sizeOfData * xy) - (x*y)
    dev1 = math.sqrt(sizeOfData * xx - x*x)
    dev2 = math.sqrt(sizeOfData * yy - y*y)
    correlation = correlation / (dev1*dev2)
    return correlation

def main():
    checkData()
    while(True):
        stockName = input("Please Enter the Stock Code: \n").upper()
        if(stockName == "EXIT"):
            sys.exit(0)
        displayData(stockName)

if __name__ == "__main__":
    main()
