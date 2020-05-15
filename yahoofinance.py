from lxml import html
import requests
import json
import pandas as pd
from StyleFrame import StyleFrame
from datetime import datetime
from yahoo_fin import stock_info as si

def get_headers():
    return {"accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
            "accept-encoding": "gzip, deflate, br",
            "accept-language": "en-GB,en;q=0.9,en-US;q=0.8,ml;q=0.7",
            "cache-control": "max-age=0",
            "dnt": "1",
            "sec-fetch-dest": "document",
            "sec-fetch-mode": "navigate",
            "sec-fetch-site": "none",
            "sec-fetch-user": "?1",
            "upgrade-insecure-requests": "1",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.122 Safari/537.36"}

#for TSE tickers remember to add 'Ticker'.TO
def getInfo(ticker):
    link = "http://finance.yahoo.com/quote/%s?p=%s" % (ticker, ticker)
    response = requests.get(link, verify=False, headers=get_headers(), timeout=10)
    parser = html.fromstring(response.text)
    summary_table = parser.xpath(
        '//div[contains(@data-test,"summary-table")]//tr')
    summary_dict = {}
    other_details_json_link = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/{0}?formatted=true&lang=en-US&region=US&modules=summaryProfile%2CfinancialData%2CrecommendationTrend%2CupgradeDowngradeHistory%2Cearnings%2CdefaultKeyStatistics%2CcalendarEvents&corsDomain=finance.yahoo.com".format(
        ticker)
    response = requests.get(other_details_json_link)
    
    loaded_summary = json.loads(response.text)

    summary = loaded_summary["quoteSummary"]["result"][0]
    y_target = summary["financialData"]["targetMeanPrice"]['raw']
    earnings_list = summary["calendarEvents"]['earnings']
    eps = summary["defaultKeyStatistics"]["trailingEps"]['raw']
    datelist = []

    #adjust the earnings date 
    for i in earnings_list['earningsDate']:
        datelist.append(i['fmt'])
        
    earnings_date = ' to '.join(datelist)
    
    
    for table_data in summary_table:
        raw_key = table_data.xpath(
            './/td[1]//text()')
        raw_val = table_data.xpath(
            './/td[2]//text()')
        key = ''.join(raw_key).strip()
        value = ''.join(raw_val).strip()
        summary_dict.update({key: value})
        
    summary_dict.update({'EPS': eps,
                         'Earnings Date': earnings_date})
    
    summary_dict["Dividend"]=summary_dict.pop("Forward Dividend & Yield")
    summary_dict["PE Ratio"]=summary_dict.pop("PE Ratio (TTM)")
    summary_dict["1 Yr Target"]=y_target
    
    del summary_dict['1y Target Est']
    del summary_dict["Bid"]
    del summary_dict["Ask"]
    del summary_dict["EPS (TTM)"] 
    
    return summary_dict

def get_price(ticker):
    return int(round(si.get_live_price(ticker)*100))/100

def formatData(input):
    input_df=pd.read_excel(input)
    #convert stock tickers into a dataframe
    df_list=[]
    ticker_list=[]
    price_list=[]
    for index, row in input_df.iterrows():
        ticker=row["Stock Names"]
        if (index==0): 
            first_input=pd.DataFrame(getInfo(ticker).items())
            t_input=first_input.T
            df_list.append(t_input)
            ticker_list.append(ticker)
        else:
            first_add_df=pd.DataFrame(getInfo(ticker).items())
            add_df=first_add_df.transpose()
            df_list.append(add_df.drop(0, axis=0))
            ticker_list.append(ticker)
    
    stock_df=pd.concat(df_list)
    
    for i in range (14):
        stock_df.rename(columns={i:stock_df.loc[0,i]}, inplace=True)
        
    stock_df=stock_df.drop(0, axis=0)
    
    for ticker in ticker_list:
        price_list.append(get_price(ticker))
        
    stock_df.insert(loc=0,column='Stock Names', value=ticker_list)
    stock_df.insert(loc=1, column='Current Price', value=price_list)
    
    return stock_df

def get_calendar(df):
    date_list=[]
    for index, row in df.iterrows():
        date=datetime.strptime(df.loc[index, 'Earnings Date'][:10], '%Y-%m-%d')
        date_list.append(date)
    
    date_list.sort()
    sorted_dates = [datetime.strftime(ts, "%Y-%m-%d") for ts in date_list]
    
    counter=-1
    for dates in sorted_dates:
        for index, row in df.iterrows():
            if (df.loc[index, 'Earnings Date'][:10]==dates):
                counter+=1
                sorted_dates[counter]=df.loc[index,'Stock Names'] + ': '+ dates


#export the data to a more formatted excel document
def styledExcel():
    stock_df = formatData('input.xlsx')
    stock_df.reset_index(drop=True,inplace=True)
    sf=StyleFrame(stock_df)
    col_list=list(stock_df.columns)
    writer = StyleFrame.ExcelWriter('C:\\Users\\matth\\Desktop\\style_version.xlsx')
    sf.to_excel(excel_writer=writer, sheet_name = 'Stock Info', best_fit=col_list)  
    writer.save()
    
styledExcel()









