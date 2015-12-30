import numpy as np
import pandas as pd
import xlrd


def read_sat_file(path):

    book = xlrd.open_workbook(path)
    data = pd.read_excel(book, header=0, engine='xlrd')
    names = ['Asset name', 'Position date & time', 'Latitude', 'Longitude', 'Average speed', 'Heading', 'Wind speed',
             'Wind direction', 'Wind & swell wave height', 'Wind wave height', 'Swell wave direction',
             'Swell wave height', 'Distance moved']
    data = data.loc[1:, names]
    data.columns = ['ship_name', 'utc_datetime', 'latitude', 'longitude', 'obs_speed', 'course', 'wind_speed',
                    'wind_dir', 'total_wave_height', 'wind_wave_height', 'swell_wave_dir', 'swell_wave_height',
                    'obs_miles']

    def knots_to_bft(knots):
        bft_thresholds = [1, 3, 6, 10, 16, 21, 27, 33, 40, 47, 55, 63, 69]
        for i in range(len(bft_thresholds)):
            if knots <= bft_thresholds[i]:
                return i
        return 12
    knots_to_bft = np.vectorize(knots_to_bft)

    data['wind_bft'] = knots_to_bft(data.wind_speed.astype('float'))

    return data


if __name__ == '__main__':

    import numpy as np

    pd.set_option('display.width', 160)
    path = 'C:\\py\\nepr\\all_vsl_sat.xlsx'
    data = read_sat_file(path)
    print(data[(data.ship_name == 'PHOENIX RISING') &
               (data.utc_datetime <= np.datetime64('2015-11-21')) &
               (data.utc_datetime >= np.datetime64('2015-11-17'))])



# # print(data.iloc[0:9, :])
# # print(data.columns)
# data = data.loc[1:, ['Asset name', 'Position date & time', 'Average speed', 'Wind speed', 'Wind & swell wave height', 'Wind wave height', 'Swell wave height']]
# # data = data.dropna()
# # delta = data['Wind & swell wave height'].astype('float') - (data['Wind wave height'].astype('float') ** 2 + data['Swell wave height'].astype('float') ** 2) ** 0.5
# # print(delta.mean())
# print(data[(data['Asset name'] == 'PHOENIX RISING') &
#            (data['Position date & time'] <= pandas.DatetimeIndex([datetime.datetime(2015, 11, 21)])[0]) &
#            (data['Position date & time'] >= pandas.DatetimeIndex([datetime.datetime(2015, 11, 17)])[0])])
# # equivalently, using numpy:
# print(data[(data['Asset name'] == 'PHOENIX RISING') &
#            (data['Position date & time'] <= numpy.datetime64('2015-11-21')) &
#            (data['Position date & time'] >= numpy.datetime64('2015-11-17'))])
# data[data['Asset name'] == 'PHOENIX RISING']['Wind speed'].astype('float').plot()
