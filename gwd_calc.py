nepr.loc[nepr.steaming != 'Yes', 'steaming'] = 0
nepr.loc[nepr.steaming == 'Yes', 'steaming'] = 1



nepr['gwd_ind'] = 0
nepr.gwd_ind = ((nepr.steaming == 1) &
                (np.max(nepr.loc[:, ['wind_wave_height', 'swell_wave_height']], axis= 1) < 1.5)
                ).astype('int')






(nepr.wind_wave_height ** 2 + nepr.swell_wave_height ** 2) ** 0.5
np.max(nepr.loc[nepr.steaming == 1, ['wind_wave_height', 'swell_wave_height']], axis= 1)
