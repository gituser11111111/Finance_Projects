
import numpy as np
import pandas as pd
import datetime as dt
import matplotlib.pyplot as plt
import pandas_datareader as pdr
from pandas_datareader import wb

# data = pd.read_excel(r'C:\Users\wbouley\Documents\Historical_stock_CSVs/Dow jones history 2.23.2022.xlsx', index_col=0, parse_dates=True)
# data.head()
# data.dtypes
# data.index
#
# data.loc['2022-02-23']
# data.loc['2021-01-01': '1990-01-01']
# data.iloc[0]
#
# type(data)
# type(data[' Close'])
# type(data[' Open'])
#
# day_change = data[' Open'] - data[' Close']
# day_change
# type(day_change)
#
# day_pct_change = (data[' Close'] - data[' Open']) /data[' Open']*100
# day_pct_change
# type(day_pct_change)
#
# data[' Close'].iloc[0]
# data[' Close'].iloc[-1]
# norm = data[' Close']/data[' Close'].iloc[-1]
# norm
#
# data['Day_%_Change'] = (data[' Close'] - data[' Open']) /data[' Open']*100
#
# data['Day_Change'] = data[' Open'] - data[' Close']
# day_change
#
# data.head()
#
# data['Normalized'] = data[' Close'] / data[' Close'].iloc[0]
#
# data.head()
#
# data[' Close'].min()
# data[' Open'].min()
# data[' Close'].argmin()
# data[' Open'].argmin()
# data.min()
# data[' Close'].max()
# data[' Open'].argmax()
# data[' Close'].max()
# data[' Close'].argmax()
# data[' Close'].mean()
# data[' Close'].median()
#
# get_ipython().run_line_magic('matplotlib', 'notebook')
#
# data.plot()
# data[' Close'].plot()
#
# fig, ax = plt.subplots()
# data[' Close'].plot(ax=ax)
# ax.set_ylabel('Price')
# ax.set_title('DJI Historical Prices')
#
# fig, ax = plt.subplots(2,2)
# data[' Open'].plot(ax=ax[0,0], title='Open')
# data[' High'].plot(ax=ax[0,1], title='High')
# data[' Low'].plot(ax=ax[1,0], title='Low')
# data[' Close'].plot(ax=ax[1,1], title='Close')
# plt.tight_layout()
#
# fig, ax = plt.subplots()
# data['Normalized'].loc['2021-07-01':'2021-06-01'].plot.bar(ax=ax)
#
# fig, ax = plt.subplots()
# data['Normalized'].loc['2021-07-01':'2021-06-01'].plot.barh(ax=ax)
#
# ticker= '^DJI'
# start= dt.datetime(1992, 1, 1)
# end = dt.datetime()
#
# data = pdr.get_data_yahoo(ticker, start)
# data.head()
# data.index
# data.dtypes
# data.tail()
# data2 = pdr.get_data_stooq(ticker, start)
# data2.head()
# data2.tail()
#
# nasdaq_sym = pdr.get_nasdaq_symbols()
# nasdaq_sym
#
# len(nasdaq_sym)
#
#
# matches = wb.search('gdp.*capita.*const')
# country_gdp = wb.download(indicator='NY.GDP.PCAP.KD', country=['US', 'CA', 'MX'], start=1990, end=2022)
# country_gdp
#
# ticker= '^DJI'
# start= dt.datetime(1992, 1, 1)
# end = dt.datetime(2022,3,2)
#
# data = pdr.get_data_yahoo(ticker, start)
# data['Log returns'] = np.log(data['Close']/data['Close'].shift())
#
# data.head()
#
# data['Close']/data['Close'].shift()
# data['Close'].shift()
# #daily standard deviation
# data['Log returns'].std()
#
# volatility = data['Log returns'].std()*252**0.5
# volatility
#
# str_vol = str(round(volatility, 4)*100)
# str_vol
#
# fig, ax = plt.subplots()
# data['Log returns'].hist(ax=ax, bins=50, alpha=0.8, color='blue')
# ax.set_xlabel('Log return')
# ax.set_ylabel('Frequency of log return')
# ax.set_title(str(ticker) + ' volatility')
#
# data.head()
#
# data['MA10'] = data['Close'].rolling(10).mean()
# data['MA25'] = data['Close'].rolling(25).mean()
#
# data.tail()
#
# data['EMA10'] = data['Close'].ewm(span=10, adjust=False).mean()
# data.tail()
#
# fig, ax = plt.subplots()
# data[['MA10', 'EMA10']].loc['2022-01-01':].plot(ax=ax)
# data['Close'].loc['2022-01-01':].plot(ax=ax, alpha=0.25)
#
# ticker= '^DJI'
# start= dt.datetime(1992, 1, 1)
# end = dt.datetime(2022,3,2)

data = pdr.get_data_yahoo(ticker, start)
data['Log returns'] = np.log(data['Close'] / data['Close'].shift())
data['MA10'] = data['Close'].rolling(10).mean()
data['MA25'] = data['Close'].rolling(25).mean()
data['EMA10'] = data['Close'].ewm(span=10, adjust=False).mean()
data.tail()

volatility = data['Log returns'].std() * 252 ** 0.5
print(volatility)
get_ipython().run_line_magic('matplotlib', 'notebook')

exp1 = data['Close'].ewm(span=12, adjust=False).mean()
exp2 = data['Close'].ewm(span=26, adjust=False).mean()

data['MACD'] = exp1 - exp2
data['Signal line'] = data['MACD'].ewm(span=9, adjust=False).mean()

data.tail()

fig, ax = plt.subplots()
data[['MACD', 'Signal line']].loc['2022-01-01':].plot(ax=ax)
data['Close'].loc['2022-01-01':].plot(ax=ax, alpha=0.25, secondary_y=True)

high14 = data['High'].rolling(14).max()
low14 = data['Low'].rolling(14).min()
data['%K'] = (data['Close'] - low14) * 100 / (high14 - low14)
data['%D'] = data['%K'].rolling(3).mean()
data.tail()

fig, ax = plt.subplots()
data[['%K', '%D']].loc['2022-01-01':].plot(ax=ax)
ax.axhline(80, c='r', alpha=0.3)
ax.axhline(20, c='r', alpha=0.3)
data['Close'].loc['2022-01-01':].plot(ax=ax, alpha=0.3, secondary_y=True)

writer = pd.ExcelWriter('DJI_Technical_Analysis',
                        engine='xlsxwriter',
                        date_format='yyyy-mm-dd',
                        datetime_format='yyy-mm-dd')
workbook = writer.book

# create a fomat for a red cell
red_cell = workbook.add_format({
    'bg_color': '#FFC7CE',
    'font_color': '9C0006'
})

# create a fomat for a green cell
green_cell = workbook.add_format({
    'bg_color': '#C6EFCE',
    'font_color': '006100'
})

# 10 day moving average sheet
#
#

sheet_name = 'MA10'
data[['Close', 'MA10']].to_excel(writer, sheet_name=sheet_name)
worksheet = writer.sheets[sheet_name]

# set column width of date
worksheet.set_column(0, 0, 15)

for col in range(1, 3):
    # create a conditional formatted of type formula
    worksheet.conditional_format(1, col, len(data), col, {
        'type': 'formula',
        'criteria': '=B2>=C2',
        'format': green_cell
    })

    worksheet.conditional_format(1, col, len(data), col, {
        'type': 'formula',
        'criteria': '=B2<C2',
        'format': red_cell
    })

# create a new chart object.
chart1 = workbook.add_chart({'type': 'line'})

# add series to the chart
chart1.add_series({
    'name': '^DJI',
    'categories': [sheet_name, 1, 0, len(data), 0],
    'values': [sheet_name, 1, 1, len(data), 1]
})

chart2 = workbook.add_chart({'type': 'line'})

chart2.add_series({
    'name': sheet_name,
    'categories': [sheet_name, 1, 0, len(data), 0],
    'values': [sheet_name, 1, 2, len(data), 2]
})

# combine charts and insert axis and title names
chart1.combine(chart2)
chart1.set_title({'name': sheet_name + ' ^DJI'})
chart1.set_x_axis({'name': 'Date'})
chart1.set_y_axis({'name': 'Price'})

# insert chart into the worksheet
worksheet.insert_chart('E2', chart1)

# MACD chart
#
#

sheet_name = 'MACD'
data[['Close', 'MACD', 'Signal line']].to_excel(writer, sheet_name=sheet_name)
worksheet = writer.sheets[sheet_name]

# set column width of Date
worksheet.set_column(0, 0, 15)

for col in range(1, 4):
    # create conditional formatted of type formula
    worksheet.conditional_format(1, col, len(data), col, {
        'type': 'formula',
        'criteria': '=C2>=D2',
        'format': green_cell
    })

    # create conditional formatted of type formula
    worksheet.conditional_format(1, col, len(data), col, {
        'type': 'formula',
        'criteria': '=C2<D2',
        'format': red_cell
    })

# create a new chart object.
chart1 = workbook.add_chart({'type': 'line'})

# add series to the chart.
chart1.add_series({
    'name': 'MACD',
    'categories': [sheet_name, 1, 0, len(data), 0],
    'values': [sheet_name, 1, 2, len(data), 2]
})

# create a new chart object
chart2 = workbook.add_chart({'type': 'line'})

# add series to the chart.
chart1.add_series({
    'name': 'Signal line',
    'categories': [sheet_name, 1, 0, len(data), 0],
    'values': [sheet_name, 1, 3, len(data), 3]
})

# combine charts and insert axis and title names
chart1.combine(chart2)
chart1.set_title({'name': sheet_name + ' ^DJI'})
chart1.set_x_axis({'name': 'Date'})
chart1.set_y_axis({'name': 'Value'})

# set labels on x axis not on 0
chart1.set_x_axis({
    'label_position': 'low',
    'num_font': {'rotation': 45}
})

# insert the chart into the worksheet
worksheet.insert_chart('F2', chart1)

# Stochastic chart
#
#

sheet_name = 'Stochastic'
data[['Close', '%K', '%D']].to_excel(writer, sheet_name=sheet_name)
worksheet = writer.sheets[sheet_name]

# set column width of Date
worksheet.set_column(0, 0, 15)

for col in range(1, 4):
    # create conditional formatted of type formula
    worksheet.conditional_format(1, col, len(data), col, {
        'type': 'formula',
        'criteria': '=C2>=D2',
        'format': green_cell
    })

    # create conditional formatted of type formula
    worksheet.conditional_format(1, col, len(data), col, {
        'type': 'formula',
        'criteria': '=C2<D2',
        'format': red_cell
    })

# create a new chart object.
chart1 = workbook.add_chart({'type': 'line'})

# add series to the chart.
chart1.add_series({
    'name': '%K',
    'categories': [sheet_name, 1, 0, len(data), 0],
    'values': [sheet_name, 1, 2, len(data), 2]
})

# create a new chart object
chart2 = workbook.add_chart({'type': 'line'})

# add series to the chart.
chart1.add_series({
    'name': '%D',
    'categories': [sheet_name, 1, 0, len(data), 0],
    'values': [sheet_name, 1, 3, len(data), 3]
})

# combine charts and insert axis and title names
chart1.combine(chart2)
chart1.set_title({'name': sheet_name + ' ^DJI'})
chart1.set_x_axis({'name': 'Date'})
chart1.set_y_axis({'name': 'Value'})

# set labels on x axis not on 0
chart1.set_x_axis({
    'label_position': 'low',
    'num_font': {'rotation': 45}
})

# insert the chart into the worksheet
worksheet.insert_chart('F2', chart1)

# writer.save()
writer.close()





