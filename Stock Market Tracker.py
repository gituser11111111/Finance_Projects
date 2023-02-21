import pandas as pd
from alpha_vantage.timeseries import TimeSeries
import time

API = 'your api key from alpha_vantage'

print('Enter ticker: ')
ticker = input('').upper()

print('Enter time interval: 1min, 5min, 15min, or 30min.')
interval = input('')

ts = TimeSeries(key=API, output_format='pandas')
data, meta_data = ts.get_intraday(symbol=str(ticker), interval=str(interval), outputsize='full')
print(data)

i = 1

# while i == 1:
#     data, meta_data = ts.get_intraday(symbol='MSFT', interval='1min', outputsize='full')
#     data.to_excel("output.xlsx")
#     time.sleep(60)

close_data = data['4. close']
percent_change = close_data.pct_change()

print(percent_change)

last_change = percent.change[-1]

if abs(last_change) > 0.0004:
    print(f'{ticker} Alert:' + last_change)

df = pd.DataFrame(data)

print(df)