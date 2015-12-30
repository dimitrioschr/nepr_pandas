nepr.loc[nepr.steaming != 'Yes', 'steaming'] = 0
nepr.loc[nepr.steaming == 'Yes', 'steaming'] = 1

nepr.hours - (nepr.utc_datetime - nepr.utc_datetime.shift()).astype('timedelta64[m]') / 60
nepr.obs_speed - (nepr.obs_miles / nepr.hours) * nepr.steaming
nepr.slip - (nepr.eng_miles - nepr.obs_miles) / nepr.eng_miles
nepr.voy_hours - np.cumsum(nepr.hours) * nepr.steaming
nepr.voy_obs_miles - np.cumsum(nepr.obs_miles)
nepr.voy_eng_miles - np.cumsum(nepr.eng_miles)
nepr.voy_speed - np.cumsum(nepr.obs_miles) / np.cumsum(nepr.hours) * nepr.steaming
nepr.voy_slip - (nepr.voy_eng_miles - nepr.voy_obs_miles) / nepr.voy_eng_miles

-(nepr.hfo_rob - nepr.hfo_rob.shift()) * (nepr.steaming) * 24 / nepr.hours # consumption expressed as MT/24HR
-(nepr.hfo_rob - nepr.hfo_rob[0]) * (nepr.steaming) * 24 / nepr.voy_hours  # voyage consumption as MT/24HR
# very easy to go to gwd-consumption from above

