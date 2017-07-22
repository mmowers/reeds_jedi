from __future__ import division
import gdxpds
import pandas as pd
import win32com.client as win32
import os
import sys

#switches
test_switch = False
state_switch = False
wbvis_switch = True

#get reeds output data
dfs = gdxpds.to_dataframes(r"\\nrelqnap01d\ReEDS\FY17-JEDI-MRM\runs\JEDI 2017-07-20\gdxfiles\JEDI.gdx")

this_dir = os.path.dirname(os.path.realpath(__file__))

df_full = dfs['Jedi']
df_full.rename(columns={'bigQ': 'tech', 'allyears': 'year', 'jedi_cat': 'cat'}, inplace=True)

#convert costs from 2004$ to 2015$
cost_cols = ['cost_capital', 'cost_om']
row_criteria = df_full['cat'].isin(cost_cols)
df_full.loc[row_criteria, 'Value'] = df_full.loc[row_criteria, 'Value'] / 0.796636801524834

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
df_full = df_full.groupby(['tech', 'st', 'year', 'cat'], as_index=False).sum()

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

#Read in workbook inputs and outputs
df_variables = pd.read_csv(this_dir + r'\inputs\variables.csv')
df_outputs = pd.read_csv(this_dir + r'\inputs\outputs.csv')
df_techs = pd.read_csv(this_dir + r'\inputs\techs.csv')

#initialize outputs
df_econ = pd.DataFrame(columns=['jedi_scenario', 'tech', 'st' ,'year' , 'type', 'econ_year', 'econ_cat', 'econ_type', 'econ_val'])

#loop through techs
for x, tech in enumerate(df_techs['tech'].values.tolist()):
    print('tech = ' + tech)
    #filter to just this tech
    df_tech = df_full[df_full['tech'] == tech]
    df_const = df_constants[df_constants['tech'] == tech]
    df_jedi_scen = df_jedi_scenarios[df_jedi_scenarios['tech'] == tech]
    df_var = df_variables[df_variables['tech']==tech]
    df_out = df_outputs[df_outputs['tech']==tech]

    project_size = df_techs[df_techs['tech'] == tech]['project_size'].iloc[0] #MW
    
    
    #first, open jedi workbook
    excel = win32.Dispatch('Excel.Application')
    if wbvis_switch:
        excel.Visible = True
    wb = excel.Workbooks.Open(this_dir + '\\jedi_models\\' + df_techs['jedi_model'][x])
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

        #Do existing plants first
        df_exist = df_scen[df_scen['cat']=='exist_retire']
        #set the inputs for existing plants
        for i, r in df_var.iterrows():
             ws_in.Range(r['cell']).Value = r['exist_val']
        #grab the economic outputs
        df_econ_outputs = pd.DataFrame(columns=['cat', 'type','val'])
        for i, r in df_out.iterrows():
            df_econ_outputs = df_econ_outputs.append({'cat': r['cat'], 'type': r['type'], 'val': float(ws_out.Range(r['cell']).Value)}, ignore_index=True)
        #Apply outputs to each existing plant
        for i, r in df_exist.iterrows():
            mult = r['Value']/project_size
            years = [i + 2010 for i in list(range(r['year'] - 2010 - 1))]
            for y in years:
                for j, e in df_econ_outputs[df_econ_outputs['cat']=='Operation'].iterrows():
                    df_econ = df_econ.append({'jedi_scenario': scen_name, 'tech': tech, 'st': r['st'], 'year': r['year'], 'type': 'exist', 'econ_year': y, 'econ_cat': e['cat'], 'econ_type': e['type'], 'econ_val': mult*e['val']}, ignore_index=True)

        #Now do new capacity
        df_new = df_scen[df_scen['cat']!='exist_retire']
        #pivot df_new so that we can iterate on rows, and add columns for cap_cost and om_cost
        df_new = df_new.pivot_table(index=['tech', 'st', 'year'], columns='cat', values='Value').reset_index()
        df_new['cap_cost'] = df_new['cost_capital']/df_new['capacity']/1000
        df_new['om_cost'] = df_new['cost_om']/df_new['capacity']/1000
        #now, loop through df rows, fill in new capital and o&m cost, and get associated economic impacts
        c = 1
        for i, r in df_new.iterrows():
            #set region as state if state_switch is True (WARNING: MAKE SURE ALL DATA IS MAPPED TO VALID STATES). Otherwise, United States will be used
            if state_switch:
                ws_in.Range('B13').Value = r['st'] #THIS NEEDS TO BE UPDATED TO USE CONSTANTS
            if c%100 == 0:
                print(str(c) + '/'+str(len(df_new)))
            for j, ro in df_var.iterrows():
                 ws_in.Range(ro['cell']).Value = r[ro['cat']]
            #grab the economic outputs
            df_econ_outputs = pd.DataFrame(columns=['cat', 'type','val'])
            for j, ro in df_out.iterrows():
                df_econ_outputs = df_econ_outputs.append({'cat': ro['cat'], 'type': ro['type'], 'val': float(ws_out.Range(ro['cell']).Value)}, ignore_index=True)
            #we need to scale the outputs by the capacity, since we used a 100 MW project to create the outputs
            mult = r['capacity']/project_size
            #Construction results: half of outputs are in solve year t, and half are in t-1
            years = [r['year'] -1, r['year']]
            for y in years:
                for j, e in df_econ_outputs[df_econ_outputs['cat']=='Construction'].iterrows():
                    df_econ = df_econ.append({'jedi_scenario': scen_name, 'tech': tech, 'st': r['st'], 'year': r['year'], 'type': 'new', 'econ_year': y, 'econ_cat': e['cat'], 'econ_type': e['type'], 'econ_val': mult*e['val']/2}, ignore_index=True)
            #Operation results: annual construction outputs start at solve year t and go across lifetime of plant.
            years = [i + r['year'] for i in list(range(df_techs['lifetime'][x]))]
            for y in years:
                for j, e in df_econ_outputs[df_econ_outputs['cat']=='Operation'].iterrows():
                    df_econ = df_econ.append({'jedi_scenario': scen_name, 'tech': tech, 'st': r['st'], 'year': r['year'], 'type': 'new', 'econ_year': y, 'econ_cat': e['cat'], 'econ_type': e['type'], 'econ_val': mult*e['val']}, ignore_index=True)
            c = c + 1
    wb.Close(False)
#drop na columns example: x = df_full.dropna(how='all', subset=['exist_retire'])
df_econ.to_csv(this_dir + r'\outputs\out.csv', index=False)

