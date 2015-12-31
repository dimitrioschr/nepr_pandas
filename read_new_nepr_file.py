import numpy as np
import pandas as pd
import xlrd


def read_new_nepr_file(path):

    book = xlrd.open_workbook(path)

    # get the top-side values
    def get_voyage_data(book):
        voyage_data = {}

        voyage_data['ship_name'] = book.sheet_by_index(2).cell(0, 7).value
        voyage_data['voyage_number'] = book.sheet_by_index(2).cell(1, 7).value
        voyage_data['load_condition'] = book.sheet_by_index(2).cell(2, 7).value
        voyage_data['cargo_desc'] = book.sheet_by_index(2).cell(3, 7).value if book.sheet_by_index(2).cell(3, 7).value else None
        voyage_data['cargo_weight'] = book.sheet_by_index(2).cell(4, 7).value if book.sheet_by_index(2).cell(4, 7).value else None

        voyage_data['departure_port'] = book.sheet_by_index(2).cell(0, 12).value
        voyage_data['departure_date'] = xlrd.xldate_as_tuple(book.sheet_by_index(2).cell(1, 12).value, book.datemode) if book.sheet_by_index(2).cell(1, 12).value else None
        if voyage_data['departure_date']:
            voyage_data['departure_date'] = np.datetime64(
                    '{:>04}'.format(str(voyage_data['departure_date'][0])) +
                    '-' +
                    '{:>02}'.format(str(voyage_data['departure_date'][1])) +
                    '-' +
                    '{:>02}'.format(str(voyage_data['departure_date'][2]))
            )
        voyage_data['departure_draft_fore'] = book.sheet_by_index(2).cell(2, 12).value
        voyage_data['departure_draft_mid'] = book.sheet_by_index(2).cell(3, 12).value
        voyage_data['departure_draft_aft'] = book.sheet_by_index(2).cell(4, 12).value

        voyage_data['arrival_port'] = book.sheet_by_index(2).cell(0, 16).value
        voyage_data['arrival_date'] = xlrd.xldate_as_tuple(book.sheet_by_index(2).cell(1, 16).value, book.datemode) if book.sheet_by_index(2).cell(1, 16).value else None
        if voyage_data['arrival_date']:
            voyage_data['arrival_date'] = np.datetime64(
                    '{:>04}'.format(str(voyage_data['arrival_date'][0])) +
                    '-' +
                    '{:>02}'.format(str(voyage_data['arrival_date'][1])) +
                    '-' +
                    '{:>02}'.format(str(voyage_data['arrival_date'][2]))
            )
        voyage_data['arrival_draft_fore'] = book.sheet_by_index(2).cell(2, 16).value
        voyage_data['arrival_draft_mid'] = book.sheet_by_index(2).cell(3, 16).value
        voyage_data['arrival_draft_aft'] = book.sheet_by_index(2).cell(4, 16).value

        voyage_data['head_charterer'] = book.sheet_by_index(2).cell(0, 20).value if book.sheet_by_index(2).cell(0, 20).value else None
        voyage_data['head_charterer_date'] = xlrd.xldate_as_tuple(book.sheet_by_index(2).cell(1, 20).value, book.datemode) if book.sheet_by_index(2).cell(1, 20).value else None
        if voyage_data['head_charterer_date']:
            voyage_data['head_charterer_date'] = np.datetime64(
                    '{:>04}'.format(str(voyage_data['head_charterer_date'][0])) +
                    '-' +
                    '{:>02}'.format(str(voyage_data['head_charterer_date'][1])) +
                    '-' +
                    '{:>02}'.format(str(voyage_data['head_charterer_date'][2]))
            )
        voyage_data['sub_charterer'] = book.sheet_by_index(2).cell(3, 20).value if book.sheet_by_index(2).cell(3, 20).value else None
        voyage_data['sub_charterer_date'] = xlrd.xldate_as_tuple(book.sheet_by_index(2).cell(4, 20).value, book.datemode) if book.sheet_by_index(2).cell(4, 20).value else None
        if voyage_data['sub_charterer_date']:
            voyage_data['sub_charterer_date'] = np.datetime64(
                    '{:>04}'.format(str(voyage_data['sub_charterer_date'][0])) +
                    '-' +
                    '{:>02}'.format(str(voyage_data['sub_charterer_date'][1])) +
                    '-' +
                    '{:>02}'.format(str(voyage_data['sub_charterer_date'][2]))
            )

        return voyage_data

    voyage_data = get_voyage_data(book)

    # get the CP speed and consumption from the hidden sheet
    def get_cp_descriptions(book):

        # create a dict of dicts: inner dicts are cp descriptions
        # outer dict is dict of ships

        # first create the keys for the inner dicts
        fields = []

        for i in ['e', 'f']:
            for j in ['b', 'l']:
                for k in ['s', 'h', 'm']:
                    fields.append(i + j + k)

        for i in ['pi', 'pw', 'x']:
            for j in ['h', 'm']:
                fields.append(i + j)

        fields.append('totb')
        fields.append('tots')

        # populate the dict with dicts from a dict-comp
        cp_descriptions = {}

        for i in range(3, 13):
            cp_descriptions[book.sheet_by_index(7).cell(i, 30).value] = \
                {fields[j]: book.sheet_by_index(7).cell(i, 31 + j).value for j in range(20)}

        return cp_descriptions

    cp_descriptions = get_cp_descriptions(book)

    # define the function that combines dates and times into datetimes:
    def make_datetime(date_vec, time_vec):
        valid_index = date_vec.dropna().index

        return pd.to_datetime(pd.to_datetime(date_vec[valid_index]).dt.date.astype('str') +
                              ' ' +
                              pd.to_datetime(time_vec[valid_index], format='%H:%M:%S').dt.time.astype('str')
                              )

    # define function that combines deg, min, letters into coordinates:
    # def make_coordinate(degree, minute, letter):
    #
    #     return '{:>3}'.format(str(int(degree))) + \
    #            u'\u00B0' + \
    #            ' ' + \
    #            '{:>02}'.format(str(int(minute))) + \
    #            "' " + \
    #            str(letter)

    def make_coordinate(degree, minute, letter):
        valid_index = degree.dropna().index

        return degree[valid_index].astype('int').astype('str').map('{:>3}'.format) + \
               u'\u00B0' + \
               ' ' + \
               minute[valid_index].astype('int').astype('str').map('{:>02}'.format) + \
               "' " + \
               letter[valid_index].astype('str')

    # make_coordinate = np.vectorize(make_coordinate)

    # def make_coordinate_with_index(degree, minute, letter):
    #     if degree.notnull().sum():
    #         valid_index = degree.dropna().index
    #
    #         return pd.Series(make_coordinate(degree.dropna(), minute.dropna(), letter.dropna()), index=valid_index)
    #
    #     return degree

    # define function that translates letters-direction into degrees-direction:
    def letters_to_dir(letter):
        all_letters = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE', 'S', 'SSW', 'SW', 'WSW', 'W', 'WNW', 'NW', 'NNW']
        all_dirs = [0, 22.5, 45, 67.5, 90, 112.5, 135, 157.5, 180, 202.5, 225, 247.5, 270, 292.5, 315, 337.5]
        converter = dict(zip(all_letters, all_dirs))

        return float(converter.get(letter, np.nan))     # use a default for undefined keys
    # letters_to_dir = np.vectorize(letters_to_dir)

    # def letters_to_dir_with_index(letter):
    #     if letter.notnull().sum():
    #         valid_index = letter.dropna().index
    #
    #         return pd.Series(letters_to_dir(letter.dropna()), index=valid_index)
    #
    #     return letter


    # get the table data
    position_data = pd.read_excel(book, sheetname=2, header=7, engine='xlrd')
    # exclude NAs in Date column, as well as final row with utility 1's from Excel
    valid_index = position_data['Local Date'].dropna().iloc[:-1].index

    # extract and rename columns for position, weather, bunkers, and sludge data
    # for position_data:
    position_data = position_data.ix[valid_index]
    position_data['local_datetime'] = make_datetime(position_data['Local Date'], position_data['Local Time'])
    position_data['utc_datetime'] = make_datetime(position_data['UTC      Date'], position_data['UTC Time'])
    position_data['latitude'] = make_coordinate(position_data['Lat - Deg.'],
                                                position_data['Lat - Min.'],
                                                position_data['   Lat -   N/S'])
    position_data['longitude'] = make_coordinate(position_data['Long - Deg.'],
                                                 position_data['Long - Min.'],
                                                 position_data['Long - E/W'])
    position_data['eta_datetime'] = make_datetime(position_data['ETA      Date'], position_data['ETA     Time'])
    position_names = ['local_datetime', 'utc_datetime', 'latitude', 'longitude', 'Course       (O)',
                      'Chrtrs Sailing Instruct.', 'RPM', 'M/E           Load', 'eta_datetime',
                      'Time between entries (hr)', 'Steaming (Yes/No)', 'Observed Miles', 'Engine    Miles',
                      'Observed    Speed (kn)', 'Slip', 'Total Steaming Time (hr)', 'Total Observed  Miles',
                      'Total Engine Miles', 'Voyage Average Speed (kn)', 'GWD Average Speed (kn)',
                      'Voyage Average Slip', 'Remarks']
    position_data = position_data[position_names]
    position_data.columns = ['local_datetime', 'utc_datetime', 'latitude', 'longitude', 'course', 'orders', 'rpm',
                             'load', 'eta_datetime', 'hours', 'steaming', 'obs_miles', 'eng_miles', 'obs_speed', 'slip',
                             'voy_hours', 'voy_obs_miles', 'voy_eng_miles', 'voy_speed', 'gwd_speed', 'voy_slip',
                             'remarks_position']

    # for weather_data:
    weather_data = pd.read_excel(book, sheetname=3, header=1, parse_cols=list(range(3, 25)), engine='xlrd').ix[valid_index]
    weather_names = ['Bft scale', 'Direction', 'Relative Dir (*)', 'Speed (kn)', 'Direction.1', 'Relative Dir (*).1',
                     'Height (m)', 'Direction.2', 'Relative Dir (*).2', 'Height (m).1', 'Direction.3',
                     'Relative Dir (*).3', 'Yes / No', 'Remarks']
    weather_data = weather_data[weather_names]
    weather_data.columns = ['wind_bft', 'wind_dir_letter', 'wind_dir_rel', 'curr_speed', 'curr_dir_letter',
                            'curr_dir_rel', 'wind_wave_height', 'wind_wave_dir_letter', 'wind_wave_dir_rel',
                            'swell_wave_height', 'swell_wave_dir_letter', 'swell_wave_dir_rel', 'gwd_ind',
                            'remarks_weather']
    weather_data['wind_dir'] = weather_data.wind_dir_letter.map(letters_to_dir)
    weather_data['curr_dir'] = weather_data.curr_dir_letter.map(letters_to_dir)
    weather_data['wind_wave_dir'] = weather_data.wind_wave_dir_letter.map(letters_to_dir)
    weather_data['swell_wave_dir'] = weather_data.swell_wave_dir_letter.map(letters_to_dir)

    # for bunkers_data:
    bunkers_data = pd.read_excel(book, sheetname=4, header=1, parse_cols=list(range(0, 28)), engine='xlrd').ix[valid_index]
    bunkers_names = ['HFO', 'LSHFO', 'MDO', 'MGO', 'LSMGO', 'M/E HS CO', 'M/E LS CO', 'M/E SYS OIL', 'G/E         OIL',
                     'Unnamed: 11', 'Total DFOC basis FM (lt/day)', 'HFO.1', 'LSHFO.1', 'GWD Average Consumption',
                     'MDO.1', 'MGO.1', 'LSMGO.1', 'CYL         OIL', 'LS CYL OIL', 'M/E SYS OIL.1', 'G/E         OIL.1',
                     'Remarks']
    bunkers_data = bunkers_data[bunkers_names]
    bunkers_data.columns = ['hfo_rob', 'lshfo_rob', 'mdo_rob', 'mgo_rob', 'lsmgo_rob', 'hs_co_rob', 'ls_co_rob',
                            'sys_o_rob', 'ge_o_rob', 'fm_read', 'fm_cons', 'hfo_cons', 'lshfo_cons', 'gwd_cons',
                            'mdo_cons', 'mgo_cons', 'lsmgo_cons', 'hs_co_cons', 'ls_co_cons', 'sys_o_cons', 'ge_o_cons',
                            'remarks_bunkers']

    # for sludge_data:
    sludges_data = pd.read_excel(book, sheetname=5, header=4, engine='xlrd').ix[valid_index]
    sludges_names = ['ROB Bilge Water (m3)', 'OWS Bilge Water Discharge (m3)', 'Bilge Water Delivered Ashore (m3)',
                     'Bilge Water Daily Production (m3/24hr)', 'ROB Oil Residues (Sludge) (m3)',
                     'Incineration of Oil Residues (m3)', 'Evaporation of Oil Residues (m3)',
                     'Oil Residues (Sludge) Delivered Ashore (m3)', 'Oil Residues (Sludge) Daily Production (m3/24hr)',
                     'Remarks']
    sludges_data = sludges_data[sludges_names]
    sludges_data.columns = ['bilges_rob', 'bilges_ows', 'bilges_ashore', 'bilges_prod', 'sludges_rob', 'sludges_incin',
                            'sludges_evap', 'sludges_ashore', 'sludges_prod', 'remarks_sludges']

    # join the table data
    nepr = position_data.join(weather_data, lsuffix='__position_', rsuffix='__weather_')
    nepr = nepr.join(bunkers_data, rsuffix='__bunker_')
    nepr = nepr.join(sludges_data, rsuffix='__sludge_')

    # incorporate voyage_data into nepr-DataFrame:
    for key in voyage_data:
        nepr[key] = voyage_data[key]

    # get applicable speed, consumption for CP performance
    nepr['cp_speed'] = np.nan
    nepr['cp_cons'] = np.nan
    nepr.loc[(nepr.steaming == 'Yes') & (nepr.load_condition == 'LADEN') & (nepr.orders == 'Full'), 'cp_speed'] = \
        cp_descriptions[voyage_data['ship_name']]['fls']
    nepr.loc[(nepr.steaming == 'Yes') & (nepr.load_condition == 'LADEN') & (nepr.orders == 'Eco'), 'cp_speed'] = \
        cp_descriptions[voyage_data['ship_name']]['els']
    nepr.loc[(nepr.steaming == 'Yes') & (nepr.load_condition == 'BALLAST') & (nepr.orders == 'Full'), 'cp_speed'] = \
        cp_descriptions[voyage_data['ship_name']]['fbs']
    nepr.loc[(nepr.steaming == 'Yes') & (nepr.load_condition == 'BALLAST') & (nepr.orders == 'Eco'), 'cp_speed'] = \
        cp_descriptions[voyage_data['ship_name']]['ebs']
    nepr.loc[(nepr.steaming == 'Yes') & (nepr.load_condition == 'LADEN') & (nepr.orders == 'Full'), 'cp_cons'] = \
        cp_descriptions[voyage_data['ship_name']]['flh']
    nepr.loc[(nepr.steaming == 'Yes') & (nepr.load_condition == 'LADEN') & (nepr.orders == 'Eco'), 'cp_cons'] = \
        cp_descriptions[voyage_data['ship_name']]['elh']
    nepr.loc[(nepr.steaming == 'Yes') & (nepr.load_condition == 'BALLAST') & (nepr.orders == 'Full'), 'cp_cons'] = \
        cp_descriptions[voyage_data['ship_name']]['fbh']
    nepr.loc[(nepr.steaming == 'Yes') & (nepr.load_condition == 'BALLAST') & (nepr.orders == 'Eco'), 'cp_cons'] = \
        cp_descriptions[voyage_data['ship_name']]['ebh']

    return nepr, voyage_data, cp_descriptions

if __name__ == '__main__':

    pd.set_option('display.width', 160)
    path = 'C:\\py\\nepr\\new_nepr.xlsx'
    nepr, voyage_data, cp_descriptions = read_new_nepr_file(path)
    print(nepr)