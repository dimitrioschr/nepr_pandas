import xlrd
import pandas
import numpy

pandas.set_option('display.width', 160)
path = 'C:\\py\\nepr\\all_vsl_sat.xlsx'
book = xlrd.open_workbook(path)

data = pandas.read_excel(book, header=0, engine='xlrd')
# print(data.iloc[0:9, :])
# print(data.columns)
data = data.loc[1:, ['Asset name', 'Position date & time', 'Average speed', 'Wind speed', 'Wind & swell wave height', 'Wind wave height', 'Swell wave height']]
# data = data.dropna()
# delta = data['Wind & swell wave height'].astype('float') - (data['Wind wave height'].astype('float') ** 2 + data['Swell wave height'].astype('float') ** 2) ** 0.5
# print(delta.mean())
print(data[(data['Asset name'] == 'PHOENIX RISING') &
           (data['Position date & time'] <= pandas.DatetimeIndex([datetime.datetime(2015, 11, 21)])[0]) &
           (data['Position date & time'] >= pandas.DatetimeIndex([datetime.datetime(2015, 11, 17)])[0])])
# equivalently, using numpy:
print(data[(data['Asset name'] == 'PHOENIX RISING') &
           (data['Position date & time'] <= numpy.datetime64('2015-11-21')) &
           (data['Position date & time'] >= numpy.datetime64('2015-11-17'))])

data[data['Asset name'] == 'PHOENIX RISING']['Wind speed'].astype('float').plot()