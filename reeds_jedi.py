from __future__ import division
import gdxpds
import pandas as pd
import openpyxl as opx

#get reeds output data
dfs = gdxpds.to_dataframes(r"C:\Users\mmowers\Projects\JEDI\reeds_jedi\gdx\JediWind.gdx")
df_cost = dfs['JediWindCost']
df_cap = dfs['JediWindBuilds']
df_cost.rename(columns={'jedi_cost_cat': 'cat', 'allyears': 'year', 'Value':'cost'}, inplace=True)
df_cap.rename(columns={'jedi_build_cat':'cat', 'allyears': 'year', 'Value':'capacity'}, inplace=True)
df_cap['cat'].replace('New', 'Capital', inplace=True)
df_cap['cat'].replace('Cumulative', 'OM', inplace=True)
df = pd.merge(left=df_cap, right=df_cost, how='outer', on=['cat','c','windtype', 'n', 'year'], sort=False)

#convert costs to 2015$
df['cost'] = df['cost']/0.796636801524834

#now turn capital and o&m costs and capacities into separate columns to bring into the same row
df_capital = df[df['cat'] == 'Capital']
df_om = df[df['cat'] == 'OM']
df = pd.merge(left=df_capital, right=df_om, how='outer', on=['c','windtype', 'n', 'year'], sort=False)
df.rename(columns={'capacity_x': 'capacity_capital', 'cost_x':'cost_capital', 'capacity_y': 'capacity_om', 'cost_y':'cost_om'}, inplace=True)

#merge with states
df_hierarchy = pd.read_csv('hierarchy.csv')
df = pd.merge(left=df, right=df_hierarchy, how='left', on=['n'], sort=False)

#limit to only onshore, and only to US, and only 2017 and after
df = df[df['windtype'] == 'wind-ons']
df = df[df['st'] != 'MEX']
df = df[df['year'] > 2016]

#group and sum
df = df.groupby(['st', 'year'], as_index=False).sum()

#now, for each row, open jedi workbook, fill in new capital and o&m cost, and get associated economic impact.
#first, open jedi workbook
wb = opx.load_workbook(filename = r"C:\Users\mmowers\Projects\JEDI\reeds_jedi\jedi_models\01D_JEDI_Land-based_Wind_Model_rel._W12.23.16.xlsm")
wb_in = wb['ProjectData']
wb_out = wb['SummaryResults']
for index, row in df.iterrows():
    if pd.notnull(row['cost_capital']):
        capital_cost = row['cost_capital']/row['capacity_capital']
        wb_in['B20'] = capital_cost
        
    if pd.notnull(row['cost_om']):
        om_cost = row['cost_om']/row['capacity_om']
        wb_in['B21'] = om_cost
    break