from __future__ import division
import gdxpds
import pandas as pd
import win32com.client as win32
import os

this_dir = os.path.dirname(os.path.realpath(__file__))

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

#merge with states
df_hierarchy = pd.read_csv('hierarchy.csv')
df = pd.merge(left=df, right=df_hierarchy, how='left', on=['n'], sort=False)

#limit to only onshore, and only to US, and only 2017 and after
df = df[df['windtype'] == 'wind-ons']
df = df[df['st'] != 'MEX']
df['year'] = df['year'].astype(int)
df = df[df['year'] > 2016]

#group and sum
df = df.groupby(['cat', 'n', 'st', 'year'], as_index=False).sum()

#add column for price ($/kW)
df['price'] = df['cost']/df['capacity']/1000


#add columns for 
out_cols = ['jobs_direct', 'jobs_indirect', 'jobs_induced',
            'earnings_direct', 'earnings_indirect', 'earnings_induced',
            'output_direct', 'output_indirect', 'output_induced',
            'value_add_direct', 'value_add_indirect', 'value_add_induced']
df = df.reindex(columns=df.columns.values.tolist() + out_cols)

#first, open jedi workbook
excel = win32.Dispatch('Excel.Application')
wb = excel.Workbooks.Open(this_dir + r'/jedi_models/01D_JEDI_Land-based_Wind_Model_rel._W12.23.16.xlsm')
ws_in = wb.Worksheets('ProjectData')
ws_out = wb.Worksheets('SummaryResults')
#now, fill in new capital and o&m cost, and get associated economic impact.
for i, r in df.iterrows():
    print(str(i) + '/'+str(len(df.index)))
    if r['cat']=='Capital':
        ws_in.Range('B20').Value = r['price']
        out_row_start = 28
    elif r['cat']=='OM':
        ws_in.Range('B21').Value = r['price']
        out_row_start = 34
    for col in range(4):
        for row in range(3):
            df.iloc[i,7 + col*3 + row] = ws_out.Cells(out_row_start + row, 2 + col).Value
wb.Close(False)
