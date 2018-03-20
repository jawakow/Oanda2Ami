import oandapy
import os
import sys
from win32com.client import Dispatch
from datetime import datetime
import tkinter
from sys import argv
import tkinter
from tkinter import *
import time

lastClose = 0
# Please edit account, token and environment before proceeding.
account = 12345
token = "qwerty"
AmiBroker = Dispatch("Broker.Application")
oanda = oandapy.API(environment="live",
                    access_token=token)


def ImportTickers():
    list = oanda.get_instruments(account)
    ticker = list.get("instruments")
    for count in range(0, len(ticker)):
        AmiBroker.Stocks.Add(ticker[count].get("instrument"))


# ImportTickers()
def Backfill():
    continous = 0
    while continous == 0:
        Qty = AmiBroker.Stocks.Count
        for i in range(0, Qty):
            inst = AmiBroker.Stocks(i).Ticker
            response = oanda.get_history(instrument=inst, count="5000", granularity="D", candleFormat="midpoint")
            prices = response.get("candles")
            for count in range(0, len(prices)):
                asking_time = prices[count].get("time")
                asking_time = asking_time.replace("-", "")
                asking_time = asking_time[:8]
                asking_time_MST = asking_time
                datetimeobject = datetime.strptime(asking_time, '%Y%m%d')
                asking_time = datetimeobject.strftime('%d/%m/%Y')
                asking_open = prices[count].get("openMid")
                asking_low = prices[count].get("lowMid")
                asking_high = prices[count].get("highMid")
                asking_close = prices[count].get("closeMid")
                ticker = AmiBroker.Stocks.Add(inst)
                quote = ticker.Quotations.Add(asking_time)
                quote.Open = asking_open
                quote.Low = asking_low
                quote.High = asking_high
                quote.Close = asking_close
                AmiBroker.RefreshAll()
                # print(asking_time,asking_open,asking_close)
                # AmiBroker.RefreshAll()


def Import():
    global daysToFill
    path = (os.path.dirname((os.path.realpath(__file__)))) + '\\' + 'Oanda.mst'
    path = path.replace('\\', '\\\\')
    file = open(path, 'w')
    Qty = AmiBroker.Stocks.Count
    for i in range(0, Qty):
        inst = AmiBroker.Stocks(i).Ticker
        response = oanda.get_history(instrument=inst, count=daysToFill, granularity="D", candleFormat="midpoint")
        prices = response.get("candles")
        for count in range(1, len(prices)):
            asking_time = prices[count].get("time")
            asking_time = asking_time.replace("-", "")
            asking_time = asking_time[:8]
            asking_time_MST = asking_time
            datetimeobject = datetime.strptime(asking_time, '%Y%m%d')
            asking_time = datetimeobject.strftime('%d/%m/%Y')
            asking_open = prices[count].get("openMid")
            # asking_open = prices[count - 1].get("closeMid")
            asking_low = prices[count].get("lowMid")
            asking_high = prices[count].get("highMid")
            asking_close = prices[count].get("closeMid")
            asking_volume = prices[count].get("volume")
            export = str(inst) + ',' + str(asking_time_MST) + ',' + str(asking_open) + ',' + str(
                asking_high) + ',' + str(asking_low) + ',' + str(asking_close) + ',' + str(asking_volume) + '\n'
            file.write(export)
            # print (export)
    file.close()
    path = (os.path.dirname((os.path.realpath(__file__)))) + '\\' + 'Oanda.mst'
    filename = path.replace('\\', '\\\\')
    # filename = 'C:\\Users\\jawakow\\oandapy-master\\Oanda.mst'
    # print(filename)
    AmiBroker.Import(0, filename, "mst.format")
    AmiBroker.RefreshAll()


def ImportCur():
    global daysToFill
    path = (os.path.dirname((os.path.realpath(__file__)))) + '\\' + 'Oanda.mst'
    path = path.replace('\\', '\\\\')
    file = open(path, 'w')
    ins = AmiBroker.ActiveDocument.Name

    response = oanda.get_history(instrument=ins, count=daysToFill, granularity="D", candleFormat="midpoint")
    prices = response.get("candles")
    for count in range(1, len(prices)):
        asking_time = prices[count].get("time")
        asking_time = asking_time.replace("-", "")
        asking_time = asking_time[:8]
        asking_time_MST = asking_time
        datetimeobject = datetime.strptime(asking_time, '%Y%m%d')
        asking_time = datetimeobject.strftime('%d/%m/%Y')
        asking_open = prices[count].get("openMid")
        # asking_open = prices[count - 1].get("closeMid")
        asking_low = prices[count].get("lowMid")
        asking_high = prices[count].get("highMid")
        asking_close = prices[count].get("closeMid")
        asking_volume = prices[count].get("volume")
        export = str(ins) + ',' + str(asking_time_MST) + ',' + str(asking_open) + ',' + str(asking_high) + ',' + str(
            asking_low) + ',' + str(asking_close) + ',' + str(asking_volume) + '\n'
        file.write(export)
        # print (export)
    file.close()
    path = (os.path.dirname((os.path.realpath(__file__)))) + '\\' + 'Oanda.mst'
    filename = path.replace('\\', '\\\\')
    # filename = 'C:\\Users\\jawakow\\oandapy-master\\Oanda.mst'
    # print(filename)
    AmiBroker.Import(0, filename, "mst.format")
    AmiBroker.RefreshAll()


def RT(lClose):
    global lastClose
    continous = 0

    inst = AmiBroker.ActiveDocument.Name
    response = oanda.get_history(instrument=inst, count="2", granularity="D", candleFormat="midpoint")
    prices = response.get("candles")

    for count in range(1, len(prices)):
        asking_time = prices[count].get("time")

        asking_time = asking_time.replace("-", "")

        asking_hhmm = asking_time[9:14]
        asking_time = asking_time[:8]
        asking_time_MST = asking_time
        datetimeobject = datetime.strptime(asking_time, '%Y%m%d')
        asking_time = datetimeobject.strftime('%d/%m/%Y')
        # asking_open = prices[count].get("openMid")
        asking_open = prices[count - 1].get("closeMid")
        asking_low = prices[count].get("lowMid")
        asking_high = prices[count].get("highMid")
        asking_close = prices[count].get("closeMid")
        if lClose != asking_close:
            ticker = AmiBroker.Stocks.Add(inst)
            quote = ticker.Quotations.Add(asking_time)
            # print(asking_time+' '+asking_hhmm)
            quote.Open = asking_open
            quote.Low = asking_low
            quote.High = asking_high
            quote.Close = asking_close
            AmiBroker.RefreshAll()
            lastClose = asking_close

            # print(asking_time,asking_open,asking_close)
            # AmiBroker.RefreshAll()


def pr():
    print("test")


# Main...        
top = tkinter.Tk()
top.title("Oanda2Ami")
CheckVar1 = IntVar()
CheckVar2 = IntVar()
C1 = Checkbutton(top, text="Real time", variable=CheckVar1, \
                 onvalue=1, offvalue=0, height=5, \
                 width=20)
C1.pack()
L = Label(top, text="Days to backfill")
L.pack()
df = StringVar()
E = Entry(top, textvariable=df)
E.pack()

df.set(5000)
daysToFill = df.get()

B0 = Button(top, text="Import tickers", command=ImportTickers)
B0.pack()
B = Button(top, text="Backfill all", command=Import)
B.pack()
B1 = Button(top, text="Backfill current", command=ImportCur)
B1.pack()

# Code to add widgets will go here...
while True:
    if CheckVar1.get() == 1:
        RT(lastClose)
    daysToFill = df.get()
    top.update_idletasks()
    top.update()
    time.sleep(0.0001)
