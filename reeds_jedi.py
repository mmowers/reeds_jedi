from __future__ import division
import gdxpds
import numpy as np
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
df_hierarchy = pd.read_csv(this_dir + r'\inputs\hierarchy.csv')
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


#add columns for jedi outputs
out_cols = ['jobs_direct', 'jobs_indirect', 'jobs_induced',
            'earnings_direct', 'earnings_indirect', 'earnings_induced',
            'output_direct', 'output_indirect', 'output_induced',
            'value_add_direct', 'value_add_indirect', 'value_add_induced']
df = df.reindex(columns=df.columns.values.tolist() + out_cols)

#first, open jedi workbook
excel = win32.Dispatch('Excel.Application')
#excel.Visible = True
wb = excel.Workbooks.Open(this_dir + r'/jedi_models/01D_JEDI_Land-based_Wind_Model_rel._W12.23.16.xlsm')
ws_in = wb.Worksheets('ProjectData')
ws_out = wb.Worksheets('SummaryResults')

#set default values
ws_in.Range('B13').Value = 'United States' #region of interest
ws_in.Range('B15').Value = 2015 #year of construction
ws_in.Range('B16').Value = 100 #project size (MW)
ws_in.Range('B17').Value = 1 #number of projects
ws_in.Range('B18').Value = 3000 #Turbine size (kW)
ws_in.Range('B22').Value = 2015 #dollar year
ws_in.Range('B24').Value = 'N' #Y to use default local shares etc, and N to not use them.

#read in array of local shares. Scenarios start in the 5th column (index=4)
df_local_share = pd.read_csv(this_dir + r'\inputs\wind_local_share.csv')
content_scenarios = df_local_share.columns.values.tolist()[4:]
df = df.reindex(columns=['content_scenario'] + df.columns.values.tolist())
df_full = pd.DataFrame()
for scen_name in content_scenarios:
    #set content_scenario column
    df['content_scenario'] = scen_name
    #fill in the local share cells
    for i, r in df_local_share.iterrows():
        ws_in.Range('E' + str(r['row'])).Value = r[scen_name]
    #now, loop through df rows, fill in new capital and o&m cost, and get associated economic impacts
    for i, r in df.iterrows():
        #set region as state, or comment out to use United States as region
        #ws_in.Range('B13').Value = r['st']
        if (i+1)%100 == 0:
            print(str(i+1) + '/'+str(len(df.index)))
        if r['cat']=='Capital':
            ws_in.Range('B20').Value = r['price']
            out_row_start = 28
        elif r['cat']=='OM':
            ws_in.Range('B21').Value = r['price']
            out_row_start = 34
        #we need to scale the outputs by the capacity, since we used a 100 MW project to create the outputs
        mult = r['capacity']/float(ws_in.Range('B16').Value)
        for col in range(4):
            for row in range(3):
                df.iloc[i,8 + col*3 + row] = mult * float(ws_out.Cells(out_row_start + row, 2 + col).Value)
    df_full = pd.concat([df_full, df], ignore_index=True)
    #clear jedi values from df
    for c in out_cols:
        df[c] = np.nan
wb.Close(False)
df = df_full
df.to_csv(this_dir + r'\outputs\out.csv', index=False)

