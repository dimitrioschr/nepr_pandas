import pandas
import xlrd

pandas.set_option('display.width', 200)

path = 'C:\\Users\\tech5\\Desktop\\PHOENIX RISING workout\\NEPR PHOENIX_RISING (2L2015) Long Beach-Tokyo.xlsx'

# load the book into python
book = xlrd.open_workbook(path)

# read the book into pandas
position_data = pandas.read_excel(book, sheetname=2, header=7, engine='xlrd')
weather_data = pandas.read_excel(book, sheetname=3, header=1, parse_cols=list(range(3, 25)), engine='xlrd')
bunker_data = pandas.read_excel(book, sheetname=4, header=1, engine='xlrd')
sludge_data = pandas.read_excel(book, sheetname=5, header=4, engine='xlrd')

# drop rows without dates
position_data = position_data.loc[position_data.Date.dropna().index, :]

# drop rows without steaming hours
position_data = position_data.loc[position_data['Steaming Time (hr)'].dropna().index, :]

# print(position_data)
weather_data = weather_data.loc[position_data.index]
# print(weather_data)
bunker_data = bunker_data.loc[position_data.index]
# print(bunker_data)
sludge_data = sludge_data.loc[position_data.index]
# print(sludge_data)


### combine dataframes into a big one
check = position_data.join(weather_data, lsuffix='__position_', rsuffix='__weather_').join(bunker_data, rsuffix='__bunker_').join(sludge_data, rsuffix='__sludge_')
# print(check.columns)

cyloil_data = check.iloc[:, [0, 1, 8, 9, 14, 65, 66]]
print(cyloil_data)
cyloil_data.iloc[:, 5].plot()



