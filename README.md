# stock-info-generator

The *stock-info-generator* is a tool for investors to gain access to information on the stocks they are interested in. Simply upload a list of stock tickers in a spreadsheet and you'll get back a spreadsheet of stock information from current price to market cap to analyst recommendations. It also gives you the option to see historical dividend and price data through tables and plots. Looking up individual stocks is easy but it can be difficult to keep track of multiple stocks. The stock-info-generator is there to help you effortlessly track all the many stocks in your portfolio and watchlist.

As an investor, keeping up to date with a company’s earnings report releases is a necessity. Earnings reports and company guidance can induce immense upside or downside that can lead to huge profits or losses. That’s why the *stock-info-generator* also gives the option of returning a separate sheet of the upcoming earnings dates for the companies you input.

With many companies struggling to maintain liquidity and bond yields so low, now more than ever, portfolios are in need of stocks with reliable and consistent dividends. That’s why the  *stock-info-generator* gives the option of returning a separate sheet containing dividend plots for the companies you’ve inputted. 

In such volatile circumstances, the  *stock-info-generator* is a tool that has helped me and will hopefully help other investors stay up to date with the stock market!

The following libraries were used:
* *pandas* to manipulate the data
* *styleframe* to format the excel output display
* *openpyxl* to plot the data 
* *lxml, json, requests* to scrape the yahoo finance page for stock information
* *yahoo_fin* for live price data
* *yahoo_finance* for historical dividends and prices

## How It Works

#### 1. Create a spreadsheet with the stock tickers for the stocks you are interested in using the following format.
<img src="https://user-images.githubusercontent.com/55144676/83339363-0eca7000-a29b-11ea-9629-ee9a5c60f26a.JPG" width="150" height="300">
*Input excel file with Stock Names as header and 10 stock tickers below*

#### 2. 




