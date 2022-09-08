import pandas as pd
from pytrends.request import TrendReq
from time import sleep

    
pytrends = TrendReq(hl='en-US', tz=360)

def get_trends_result(sheet_name,keyWords,dates):
    
    for i in range(len(keyWords)):
        sleep(1)
        keyword = keyWords[i]
        date = dates[i]
        #start date 3 months before date and end date 3 months after date
        start_date = date - pd.DateOffset(months=3)
        end_date = date + pd.DateOffset(months=3)
        start_date = start_date.strftime('%Y-%m-%d')
        end_date = end_date.strftime('%Y-%m-%d')
        #search interest over time with start date and end date and firmName
        pytrends.build_payload([keyword], cat=0, timeframe=f'{start_date} {end_date}')
        #get interest over time
        interest_over_time = pytrends.interest_over_time()
        trends_data = pd.DataFrame(interest_over_time)
        if trends_data.empty:
            data = pd.DataFrame()
            pass
        else:
            trends_data.drop(columns=['isPartial'], inplace=True)
            data = trends_data.T
        if(i == 0):
            all_data = data
        else:
            all_data = pd.concat([all_data, data])

    with pd.ExcelWriter('GT_data.xlsx', mode='a', if_sheet_exists='overlay') as writer:
        all_data.to_excel(writer, sheet_name=sheet_name)
    


if __name__ == '__main__':
    #pandas read xlsx file
    df_sheet_us = pd.read_excel('GT_data.xlsx', sheet_name='US')
    df_sheet_china = pd.read_excel('GT_data.xlsx', sheet_name='China')

    firmNames_us = df_sheet_us['firm_name'].to_list()
    dates_us = df_sheet_us['firm_date'].to_list()
    tickers_us = df_sheet_us['Ticker'].to_list()

    firmNames_china = df_sheet_china['firm_name'].to_list()
    dates_china = df_sheet_china['firm_date'].to_list()
    tickers_china = df_sheet_china['Ticker'].to_list()

    get_trends_result(sheet_name='US_name_overview',keyWords=firmNames_us,dates=dates_us)
    get_trends_result(sheet_name='US_ticker_overview',keyWords=tickers_us,dates=dates_us)
    # get_trends_result(sheet_name='China_name_overview',keyWords=firmNames_china,dates=dates_china)
    # get_trends_result(sheet_name='China_ticker_overview',keyWords=tickers_china,dates=dates_china)
    print("done")