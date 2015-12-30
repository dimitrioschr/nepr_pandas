import read_sat_file
import read_new_nepr_file

import pandas as pd
import numpy as np
pd.set_option('display.width', 160)


sat = read_sat_file.read_sat_file('C:\\py\\nepr\\all_vsl_sat.xlsx')
nepr, voyage_data, cp_descriptions = read_new_nepr_file.read_new_nepr_file('C:\\py\\nepr\\new_nepr2.xlsx')

sat = sat[(sat.ship_name == voyage_data['ship_name']) &
          (sat.utc_datetime >= nepr.utc_datetime.iloc[0]) &
          (sat.utc_datetime <= nepr.utc_datetime.iloc[-1])]

combo = nepr.append(sat)
combo = combo.sort_values('utc_datetime')   # combo for presentation purposes only (so far)

nepr.set_index('utc_datetime', inplace=True)
sat.set_index('utc_datetime', inplace=True)

nepr.wind_bft.plot()
sat.wind_bft.plot()

nepr.curr_speed.plot()

nepr.wind_wave_height.plot()
nepr.swell_wave_height.plot()

sat.loc[:, ['total_wave_height', 'wind_wave_height', 'swell_wave_height']].fillna(axis=0, method='ffill', inplace=True)
sat.loc[:, ['total_wave_height', 'wind_wave_height', 'swell_wave_height']].fillna(axis=1, method='ffill', inplace=True)
sat.total_wave_height.astype('float').plot()
sat.wind_wave_height.astype('float').plot()
sat.swell_wave_height.astype('float').plot()
sat.obs_speed.astype('float').plot()

nepr.wind_wave_height.plot()
nepr.swell_wave_height.plot()
nepr.obs_speed.plot()

####################################################

import matplotlib.style
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

matplotlib.style.use('ggplot')

plt.close('all')
fig, ax = plt.subplots(1)
ax.plot(nepr.index, nepr.wind_wave_height, marker='o', label='nepr wind wave height')
ax.plot(sat.index, sat.wind_wave_height.astype('float'), marker='o', label='sat wind wave height')
ax.legend(loc='upper right').get_frame().set_alpha(0.5)
fig.autofmt_xdate()

xfmt = mdates.DateFormatter('%b %d - %H:%M')
ax.xaxis.set_major_formatter(xfmt)
plt.ylim([0, plt.ylim()[1]])    # change only lower y - limit