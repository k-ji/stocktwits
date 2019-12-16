import requests
import pandas as pd
import json
import os
import shutil
import datetime
import win32com.client
rom iexfinance.stocks import Stock
from iexfinance.stocks import get_earnings_today
import time
import requests
from bs4 import BeautifulSoup
import progressbar
import urllib.request as req
import urllib.request
from datetime import datetime


# get the data using StockTwits API
def get_twits(ticker):
    # url = "https://api.stocktwits.com/api/2/streams/symbol/{0}.json".format(ticker)
    url = "https://api.stocktwits.com/api/2/streams/trending.json"
    response = requests.get(url).json()
    return response


# loops through to get data for each ticker in the tickers list
def get_twits_list(tickers):
    ret = {}
    for ticker in tickers:
        print("Getting data for", ticker)
        # error handling
        try:
            data = get_twits(ticker)
            symbol = data['symbol']['symbol']
            msgs = data['messages']
            ret.update({symbol: msgs})
        except Exception as e:
            print(e)
            print("Error getting", ticker)
    return ret


def read_tickers():
    print("Reading tickers from \"tickers.txt\":")
    f = open("tickers.txt", 'r')
    names = []
    # read tickers from tickers.txt
    for line in f:
        line = line.strip('\n')
        line = line.upper()
        line = line.strip('\t')
        names.append(line)
    print(names)
    return names


# Get list of trending symbols from stocktwits
def get_trending_symbols():
    """
    Returns:
      A dataframe with symbols and details.
    """

    url = "https://api.stocktwits.com/api/2/trending/symbols.json"
    response = requests.get(url).json()
    response['symbols'][0]

    symbols = [[twit['symbol'], twit['watchlist_count'], twit['id'], twit['title'],
                twit['aliases']] for twit in response['symbols']]

    now = datetime.now()
    time = now.strftime("%H:%M:%S")
    today = now.strftime("%Y%m%d")

    symbols = pd.DataFrame(symbols, columns=['symbol', 'watchlist_count', 'id',
                                             'title', 'aliases'])
    symbols['date'] = today
    symbols['time'] = time

    cols = ['date', 'time', 'symbol', 'watchlist_count', 'id', 'title', 'aliases']
    symbols = symbols[cols]

    return symbols


def write_to_csv(file, data):
    now = datetime.now()
    print("Name of the file: ", file)

    try:
        if os.path.exists(file):
            df = pd.read_csv(file)
            df = df.append(data)
            df.to_csv(file, header=True, index=False)
        else:
            data.to_csv(file, header=True, index=False)
        print(now.strftime("%Y-%m-%d" + " - updated trending symbols file"))
    except:
        print(now.strftime("%Y-%m-%d" + " - File is not updated. Check if file is open!"))
        

def send_email(subject,html_body,receiver):
    
    olMailItem = 0x0
    olFormatHTML = 2
    olFormatPlain = 1
    olFormatRichText = 3
    olFormatUnspecified = 0
    olMailItem = 0x0
    
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    
    newMail.BodyFormat = olFormatHTML 
    
    newMail.To = receiver
    newMail.Subject = subject
    newMail.HTMLBody = html_body
    
    newMail.Send()
    
    return
    
    
def get_iex_price_quote(syms):
    stock = Stock(syms[0], output_format='pandas', token="sk_4127ca1b4b274a01a41706b12e3d69c7")
    stock_quote = stock.get_quote()
    quote = stock_quote.loc[['latestPrice', 'extendedPrice',  'previousClose', 'changePercent', 'change',
                             'volume', 'avgTotalVolume', 'marketCap']].T
    if len(syms) > 1:

        for sym in syms[1:]:
            try:
                stock = Stock(sym, output_format='pandas', token="sk_4127ca1b4b274a01a41706b12e3d69c7")
                stock_quote = stock.get_quote()
                stock_quote = stock_quote.loc[
                    ['latestPrice', 'extendedPrice', 'previousClose', 'changePercent', 'change',
                     'volume', 'avgTotalVolume', 'marketCap']].T
            except:
                data = ["-999", "-999", "-999", "-999", "-999", "-999", "-999", "-999"]
                cols = ['latestPrice', 'extendedPrice', 'previousClose', 'changePercent', 'change',
                        'volume', 'avgTotalVolume', 'marketCap']
                stock_quote = pd.DataFrame(data).T
                stock_quote.columns = cols

            quote = quote.append(stock_quote)

    return quote


def get_finviz_stock_chart(syms, fn='./data/finviz/',period='d'):
    
    """
    Params:
        syms:   list of symbols
        fn:     file location to save images
        period: chart time frame, i.e. d/w/m
    """
    
    if type(syms) != list and not isinstance(syms, pd.Series):
        print ("Invalid input: symbol list")
        return 
    
    if len(syms) == 0:
        print("No symbol provided")
        return

    img_names = []
    
    for sym in syms:
        if period == 'd':
            sym_url = "https://finviz.com/quote.ashx?t=" + sym
        elif period == 'w':
            sym_url = "https://finviz.com/quote.ashx?t=" + sym + "&ty=c&ta=0&p=w&b=1"
        elif period == 'm':
            sym_url = "https://finviz.com/quote.ashx?t=" + sym + "&ty=c&ta=0&p=m&b=1"
        else:
            print("Invalid period. Default to daily")
            period = 'd'

        req = requests.get(sym_url)
        soup = BeautifulSoup(req.content, 'html.parser')

        chart = soup.find_all("img",id="chart0")
    
        img_url = "https://finviz.com/" + chart[0]['src']
        img_name = fn + datetime.now().strftime("%Y%m%d_%H%M_") + sym+ "" + ".jpg"
    
        #print(img_name)
        print(img_url)
    
        urllib.request.urlretrieve(img_url, img_name)
        
        img_names.append(img_name)
        

    return img_names
    
    
    # windows
dir = "C:/Users/kunji/Google Drive/Trading2019/PythonScripts/StockTwitsAPI/"
# dir = "/mnt/c/Users/kunji/Google Drive/Trading2019/PythonScripts/StockTwitsAPI/"
# mac
# dir = "/Users/kun/Google Drive/Trading2019/PythonScripts/StockTwitsAPI/"
os.chdir(dir)

# update historical trending file
file = dir + "trending_symbols.csv"
symbols = get_trending_symbols()
symbols.to_csv("lastest_trending_symbols.csv", header=True)
write_to_csv(file, symbols)

df = symbols.loc[0:9]
df = df[['date', 'time', 'symbol', 'title']]

iex_data = get_iex_price_quote(df['symbol'])
df = df.join(iex_data.reset_index(drop=True))
df['marketCap'] = [int(int(x) / 1000000) for x in df['marketCap']]

images = get_finviz_stock_chart(df['symbol'],fn=dir+'charts/')

# Send email
# symbols = symbols[['date', 'time', 'symbol', 'title']]
subject = 'Trending Symbols'
receiver = 'elainellxie@gmail.com;kun.ji.info@gmail.com'
html1 = df.to_html()
body = html1

for image in images:
    body = body + "<br><img src=" + image + "> "

send_email(subject, body, receiver)
