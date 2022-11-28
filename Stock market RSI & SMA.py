import pandas as pd
import alpha_vantage
from alpha_vantage.techindicators import TechIndicators
import matplotlib.pyplot as plt
import datetime

api = 'your api key from alpha_vantage'

period = 60

while True:
    print('Enter Ticker: ')
    ticker = input('').upper()

    print('Enter Time Interval: 1min, 5min, 15min, or 30min.')
    interval = input('')

    ti = TechIndicators(key=api, output_format='pandas')
    data_ti, meta_data = ti.get_rsi(symbol=str(ticker), interval=str(interval),
                                    time_period=period, series_type='close')
    data_sma, meta_data_sma = ti.get_sma(symbol=str(ticker), interval=str(interval),
                                         time_period=period, series_type='close')
    df1 = data_sma.iloc[1::]
    df2 = data_ti
    df1.index = df2.index

    fig, ax1 = plt.subplots(figsize=(12,8))
    ax1.plot(df1, 'b-')
    ax2 = ax1.twinx()
    ax2.plot(df2, 'r.', label='RSI')
    plt.title('SMA & RSI Graph for ' + str(ticker) + '.')
    ax1.set_ylabel('Price', loc='center')
    ax1.set_xlabel('Date', loc='center')
    plt.legend()
    plt.grid()
    plt.xticks(rotation=30)
    plt.show()
