from lxml import html
import requests
import json
import pandas as pd
from pandas.tseries.holiday import USFederalHolidayCalendar
from pandas.tseries.offsets import CustomBusinessDay
from decimal import Decimal
import yfinance as yf
from StyleFrame import StyleFrame, Styler, utils
from datetime import datetime, date
from yahoo_fin import stock_info as si
import openpyxl 
from openpyxl.chart import LineChart,Reference
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def get_headers():
    return {"accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
            "accept-encoding": "gzip, deflate, br",
            "accept-language": "en",
            "cache-control": "max-age=0",
            "dnt": "1",
            "sec-fetch-dest": "document",
            "sec-fetch-mode": "navigate",
            "sec-fetch-site": "none",
            "sec-fetch-user": "?1",
            "upgrade-insecure-requests": "1",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.122 Safari/537.36"}


#for TSE tickers remember to add '.TO' as in 'Ticker.TO'
def getInfo(ticker):
    link = "http://finance.yahoo.com/quote/%s?p=%s" % (ticker, ticker)
    response = requests.get(link, verify=False, headers=get_headers(), timeout=100)
    parser = html.fromstring(response.text)
    summary_table = parser.xpath(
        '//div[contains(@data-test,"summary-table")]//tr')
    summary_dict = {}
    
    other_details_json_link = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/{0}?formatted=true&lang=en-US&region=US&modules=summaryProfile%2CfinancialData%2CrecommendationTrend%2CupgradeDowngradeHistory%2Cearnings%2CdefaultKeyStatistics%2CcalendarEvents&corsDomain=finance.yahoo.com".format(
        ticker)
    json_response = requests.get(other_details_json_link)
    loaded_summary = json.loads(json_response.text)

    #extract information from loaded_summary dictionary
    summary = loaded_summary["quoteSummary"]["result"][0]
    y_target = summary["financialData"]["targetMeanPrice"]['raw']
    recommendation = summary["financialData"]["recommendationKey"].capitalize()
    num_of_analysts = summary["financialData"]["numberOfAnalystOpinions"]['raw']
    earnings_list = summary["calendarEvents"]['earnings']
    eps = summary["defaultKeyStatistics"]["trailingEps"]['raw']
    datelist = []
    
    #for date range if no specific earnings date reported 
    for i in earnings_list['earningsDate']:
        datelist.append(i['fmt'])
        
    earnings_date = ' to '.join(datelist)
    
    #get information from yahoo's summary table
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
    summary_dict['Number of Analysts']=num_of_analysts
    summary_dict['Analyst Recommendation']=recommendation

    del summary_dict['1y Target Est']
    del summary_dict["Bid"]
    del summary_dict["Ask"]
    del summary_dict["EPS (TTM)"] 

    return summary_dict

#get the current stock price
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
    
    for i in range (16):
        stock_df.rename(columns={i:stock_df.loc[0,i]}, inplace=True)
        
    stock_df=stock_df.drop(0, axis=0)
    
    for ticker in ticker_list:
        price_list.append(get_price(ticker))
        
    stock_df.insert(loc=0,column='Stock Names', value=ticker_list)
    stock_df.insert(loc=1, column='Current Price', value=price_list)
    stock_df.reset_index(drop=True,inplace=True)
    stock_df.replace('N/A (N/A)', 'N/A', inplace=True)
    stock_df.replace('', 'N/A', inplace=True)
    
    return stock_df, ticker_list

#returns set of chronologically ordered earnings dates to spreadsheet
def get_earnings_dates(df):
    date_list=[]
    na_dict={}
    for index, row in df.iterrows():
        try:
            dates=datetime.strptime(df.loc[index, 'Earnings Date'][:10], '%Y-%m-%d')
            date_list.append(dates)
        except:
            na_dict[df.loc[index,'Stock Names']]= 'No reported date'
    
    date_list.sort()
    sorted_dates = [datetime.strftime(ts, "%Y-%m-%d") for ts in date_list]
    sorted_dict={}
    
    for dates in sorted_dates:
        for index, row in df.iterrows():
            if (df.loc[index, 'Earnings Date'][:10]==dates):
                sorted_dict[df.loc[index,'Stock Names']] = dates + ' ' + df.loc[index, 'Earnings Date'][11:]
             
    if (na_dict):
        sorted_dict.update(na_dict)
    
    sorted_df=pd.DataFrame(list(sorted_dict.items()), columns=['Stock Names','Earnings Date'])

    return sorted_df

#export the data to a more formatted excel document
def styled_excel(file_input, dividends, prices):
    path = "C:\\Users\\matth\\Desktop\\style_version.xlsx"
    i=True
    if (dividends==True):
        while i==True:
            try:
                quarters_back=int(input('How many quarters back of dividends of dividend data would you like?\n'))
                i=False
            except:
                print("Please enter an integer")
        
    #turn the info and earnings dataframes into spreadsheets
    stock_df, ticker_list = formatData(file_input)
    sf=StyleFrame(stock_df)
    col_list=list(stock_df.columns)
    writer = StyleFrame.ExcelWriter(path)
    sf.to_excel(excel_writer=writer, sheet_name = 'Stock Info', best_fit=col_list)  
    earnings_df=get_earnings_dates(stock_df)
    sf_earnings=StyleFrame(earnings_df)
    col_list_earnings=list(earnings_df.columns)

    sf_earnings.to_excel(excel_writer=writer, sheet_name = 'Upcoming Earnings Dates', best_fit=col_list_earnings)  
    
    if (dividends==True):
        for ticker in ticker_list:
            get_dividends(ticker, writer, quarters_back)
    
    if (prices==True):
            start_in=input("Start date (YYYY-MM-DD): ")
            end_in=input("End date (YYYY-MM-DD): ")
            start = (pd.to_datetime(start_in)+pd.tseries.offsets.BusinessDay(n = 0)).strftime("%Y-%m-%d")
            end = (pd.to_datetime(end_in)+pd.tseries.offsets.BusinessDay(n = 1)).strftime("%Y-%m-%d")
            
            for ticker in ticker_list:
                get_prices(ticker, writer, start, end)
    
    writer.save()
    
    #plotting dividends and historical prices
    if (dividends==True):
        plot_dividends(ticker_list, path, quarters_back)
    
    if (prices==True):
        plot_prices(ticker_list, path, start, end)


def get_dividends(ticker, writer, quarters_back):
    stock = yf.Ticker(ticker)
    div_df=stock.actions
    if (len(div_df)==0 or (div_df.loc[div_df.index[0], 'Dividends']==0 and len(div_df)==1)):
        return
        
    del div_df['Stock Splits']
    quarters_back_calc=len(div_df.index)-quarters_back
    
    kept_div_df=div_df.drop(div_df.index[list(range(quarters_back_calc))])
    kept_div_df.insert(loc=0, column='Ex-Dates', value=[datetime.strftime(d, "%Y-%m-%d") for d in kept_div_df.index])
    
    sf_div=StyleFrame(kept_div_df)
    sf_div.apply_column_style(cols_to_style='Ex-Dates', styler_obj=Styler(number_format=utils.number_formats.date))
    col_list_div=list(kept_div_df.columns)

    sf_div.to_excel(excel_writer=writer, sheet_name = ticker + ' Dividends', best_fit=col_list_div) 
    
#plot dividends 
def plot_dividends(ticker_list, path, quarters_back):
    for ticker in ticker_list:
        stock = yf.Ticker(ticker)
        div_df=stock.actions
        quarters_back_calc=len(div_df.index)-quarters_back
        kept_div_df=div_df.drop(div_df.index[list(range(quarters_back_calc))])
        min_val=kept_div_df['Dividends'].min()
        
        wb = openpyxl.load_workbook(path) 
        try:
            sheet = wb[ticker+ ' Dividends']
            values = Reference(sheet, min_col = 2, min_row = 2, 
                                     max_col = 2, max_row = quarters_back+1) 
            categories= Reference(sheet, min_col = 1, min_row = 2, 
                                     max_col = 1, max_row = quarters_back+1)
            sheet.sheet_properties.tabColor = '00FFFF'
            #create line chart to plot dividends
            chart = LineChart() 
            chart.legend=None
            chart.y_axis.scaling.min =int((min_val-min_val*0.5)*100)/100
            chart.add_data(values) 
            chart.set_categories(categories)
            chart.title = ticker + ' Dividends'
            chart.x_axis.title = "Dates"
            chart.y_axis.title = "Amount ($)"
            sheet.add_chart(chart, "E2") 
            wb.save(path)
            
        except:
            pass
        
def get_prices(ticker, writer, start_date, end_date):

    df=yf.download(ticker, start=start_date[:10], end=end_date[:10])
    
    i=0
    for val in df['Close']:
        df.at[df.index[i], 'Close']=int(val*100)/100
        i+=1
    
    df.insert(loc=0, column='Date', value=[datetime.strftime(d, "%Y-%m-%d") for d in df.index])
    
    del df['Open']
    del df['High']
    del df['Adj Close']
    del df['Low']
    del df['Volume']
    
    sf=StyleFrame(df)
    sf.apply_column_style(cols_to_style='Date', styler_obj=Styler(number_format=utils.number_formats.date))
    col_list=list(df.columns)

    sf.to_excel(excel_writer=writer, sheet_name = ticker+' Price', best_fit=col_list)  

#plot historical prices
def plot_prices(ticker_list, path, input_start, input_end):
    for ticker in ticker_list:
        df=yf.download(ticker, start=input_start, end=input_end)
        min_val=df['Close'].min()
        
        # days = np.busday_count(startd, endd)
        us_bd = CustomBusinessDay(calendar=USFederalHolidayCalendar())
        days=len(pd.date_range(start=input_start,end=input_end, freq=us_bd))

        wb = openpyxl.load_workbook(path) 

        sheet = wb[ticker+' Price']
        sheet.sheet_properties.tabColor = 'EE82EE'
        values = Reference(sheet, min_col = 2, min_row = 2, 
                                 max_col = 2, max_row = days+1) 
        categories= Reference(sheet, min_col = 1, min_row = 2, 
                                 max_col = 1, max_row = days+1)
        
        #create line chart to plot dividends
        chart = LineChart() 
        chart.legend=None
        chart.y_axis.scaling.min = int(min_val-min_val*0.1)
        chart.add_data(values) 
        chart.set_categories(categories)
        chart.title = ticker + ' Historical Prices'
        chart.x_axis.title = "Dates"
        chart.y_axis.title = "Price ($)"
        sheet.add_chart(chart, "F2") 

        wb.save(path)
    
#just change the input excel file name to the name of your input excel file 
#remember to format the spreadsheet with "Stock Names" in cell A1 and stock tickers underneath in column A

#file_name: file_name.xlsx as string
#dividends: boolean for whether dividend history and plot is desired
#-will be prompted for number of quarters of dividend date desired 
#prices: boolean for whether historical prices and plot is desired
#-will be prompted for date range if prices==True
#styled_excel(file_name.xlsx, dividends, prices)
#example here


styled_excel('input.xlsx', True, True)













