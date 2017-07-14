from __future__ import division
import gdxpds
import pandas as pd
import win32com.client as win32
import os
import sys

#switches
test_switch = False
state_switch = False
wbvis_switch = False

#get reeds output data
dfs = gdxpds.to_dataframes(r"\\nrelqnap01d\ReEDS\FY17-JEDI-MRM\runs\JEDI\gdxfiles\JEDI.gdx")

this_dir = os.path.dirname(os.path.realpath(__file__))

df_cost = dfs['JediCost']
df_cap = dfs['JediCapacity']
df_cost.rename(columns={'jedi_cat': 'cat', 'bigQ': 'tech', 'allyears': 'year', 'Value':'cost'}, inplace=True)
df_cap.rename(columns={'jedi_cat': 'cat', 'bigQ': 'tech', 'allyears': 'year', 'Value':'capacity'}, inplace=True)
df_full = pd.merge(left=df_cap, right=df_cost, how='outer', on=['cat','tech', 'n', 'year'], sort=False)

#convert costs to 2015$
df_full['cost'] = df_full['cost']/0.796636801524834

#merge with states
df_hierarchy = pd.read_csv(this_dir + r'\inputs\hierarchy.csv')
df_full = pd.merge(left=df_full, right=df_hierarchy, how='left', on=['n'], sort=False)

#throw error if region is unmapped
df_unmapped = df_full[pd.isnull(df_full['st'])]
if not df_unmapped.empty:
    print(df_unmapped)
    sys.exit("unmapped regions shown above!")

#limit to only 2017 and after, and only US
df_full = df_full[df_full['st'] != 'MEXICO']
df_full['year'] = df_full['year'].astype(int)
df_full = df_full[df_full['year'] > 2016]

#Test filter
if test_switch:
    df_full = df_full[df_full['st'] == 'IOWA']

#group and sum. Aggregate n to state in the process
df_full = df_full.groupby(['cat', 'tech', 'st', 'year'], as_index=False).sum()

#add column for price ($/kW)
df_full['price'] = df_full['cost']/df_full['capacity']/1000

#add columns for jedi outputs
out_cols = ['jobs_direct', 'jobs_indirect', 'jobs_induced',
            'earnings_direct', 'earnings_indirect', 'earnings_induced',
            'output_direct', 'output_indirect', 'output_induced',
            'value_add_direct', 'value_add_indirect', 'value_add_induced']
df_full = df_full.reindex(columns=df_full.columns.values.tolist() + out_cols)

#read in constants for each tech
df_constants = pd.read_csv(this_dir + r'\inputs\constants.csv')

#read in jedi_scenarios and concatenate dataframe for each one
df_jedi_scenarios = pd.read_csv(this_dir + r'\inputs\jedi_scenarios.csv')
#read in array of local shares. Scenarios start in the 6th column (index=5)
jedi_scenarios = df_jedi_scenarios.columns.values.tolist()[5:]
df_full = df_full.reindex(columns=['jedi_scenario'] + df_full.columns.values.tolist())
df_temp = pd.DataFrame() #temporary
for scen_name in jedi_scenarios:
    df_full['jedi_scenario'] = scen_name
    df_temp = pd.concat([df_temp, df_full], ignore_index=True)
df_full = df_temp

#Read in list of jedi models
df_models = pd.read_csv(this_dir + r'\inputs\models.csv')
jedi_models = df_models.to_dict('list')
#jedi_models = dict(zip(list(df_models['tech']), list(df_models['model'])))

#Read in workbook inputs and outputs
df_variables = pd.read_csv(this_dir + r'\inputs\variables.csv')
df_outputs = pd.read_csv(this_dir + r'\inputs\outputs.csv')
df_size = pd.read_csv(this_dir + r'\inputs\project_size.csv')

#loop through techs
for x, tech in enumerate(jedi_models['tech']):
    print('tech = ' + tech)
    #filter to just this tech
    df_tech = df_full[df_full['tech'] == tech]
    df_const = df_constants[df_constants['tech'] == tech]
    df_jedi_scen = df_jedi_scenarios[df_jedi_scenarios['tech'] == tech]
    df_var = df_variables[df_variables['tech']==tech]
    df_out = df_outputs[df_outputs['tech']==tech]

    project_size = df_size[df_size['tech'] == tech]['MW'].iloc[0] #MW
    
    #first, open jedi workbook
    excel = win32.Dispatch('Excel.Application')
    if wbvis_switch:
        excel.Visible = True
    wb = excel.Workbooks.Open(this_dir + '\\jedi_models\\' + jedi_models['model'][x])
    ws_in = wb.Worksheets('ProjectData')
    ws_out = wb.Worksheets('SummaryResults')

    #set constants
    for i, r in df_const.iterrows():
        ws_in.Range(r['cell']).Value = r['value']
    for scen_name in jedi_scenarios:
        print('scenario = ' + scen_name)
        #filter to correct jedi scenario
        df_scen = df_tech[df_tech['jedi_scenario'] == scen_name]
        #fill in the local share cells
        for i, r in df_jedi_scen.iterrows():
            ws_in.Range(r['cell']).Value = r[scen_name]
        #now, loop through df rows, fill in new capital and o&m cost, and get associated economic impacts
        c = 1
        for i, r in df_scen.iterrows():
            #set region as state if state_switch is True (WARNING: MAKE SURE ALL DATA IS MAPPED TO VALID STATES). Otherwise, United States will be used
            if state_switch:
                ws_in.Range('B13').Value = r['st']
            if c%100 == 0:
                print(str(c) + '/'+str(len(df_scen)))
            cell = df_var[df_var['cat'] == r['cat']]['cell'].iloc[0]
            ws_in.Range(cell).Value = r['price']
            #we need to scale the outputs by the capacity, since we used a 100 MW project to create the outputs
            mult = r['capacity']/project_size
            for j,ro in df_out[df_out['cat'] == r['cat']].iterrows():
                df_full.loc[i, ro['type']] = mult*float(ws_out.Range(ro['cell']).Value)
            c = c + 1
    wb.Close(False)
df_full.to_csv(this_dir + r'\outputs\out.csv', index=False)

