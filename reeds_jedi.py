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

this_dir = os.path.dirname(os.path.realpath(__file__))

#get reeds output data
dfs = gdxpds.to_dataframes(r"\\nrelqnap01d\ReEDS\FY17-JEDI-MRM\runs\JEDI 2017-07-20\gdxfiles\JEDI.gdx")

#Read in workbook input csvs
df_techs = pd.read_csv(this_dir + r'\inputs\techs.csv')
df_hierarchy = pd.read_csv(this_dir + r'\inputs\hierarchy.csv')
df_constants = pd.read_csv(this_dir + r'\inputs\constants.csv')
df_jedi_scenarios = pd.read_csv(this_dir + r'\inputs\jedi_scenarios.csv')
df_variables = pd.read_csv(this_dir + r'\inputs\variables.csv')
df_outputs = pd.read_csv(this_dir + r'\inputs\outputs.csv')
df_output_cat = pd.read_csv(this_dir + r'\inputs\output_categories.csv')

df_full = dfs['Jedi']
df_full.rename(columns={'bigQ': 'tech', 'allyears': 'year', 'jedi_cat': 'cat'}, inplace=True)

#convert costs from 2004$ to 2015$
cost_cols = ['cost_capital', 'cost_om']
row_criteria = df_full['cat'].isin(cost_cols)
df_full.loc[row_criteria, 'Value'] = df_full.loc[row_criteria, 'Value'] / 0.796636801524834

#merge with states
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

#concatenate dataframe for each jedi scenario
#read in array of local shares. Scenarios start in the 6th column (index=5)
jedi_scenarios = df_jedi_scenarios.columns.values.tolist()[5:]
df_full = df_full.reindex(columns=['jedi_scenario'] + df_full.columns.values.tolist())
df_temp = pd.DataFrame() #temporary
for scen_name in jedi_scenarios:
    df_full['jedi_scenario'] = scen_name
    df_temp = pd.concat([df_temp, df_full], ignore_index=True)
df_full = df_temp

#Make another column to distinguish new capacity from existing capacity that is retiring.
df_full['type'] = 'new'
df_full.loc[df_full['cat']=='exist_retire', 'type'] = 'exist_retire'
df_full.loc[df_full['type']=='exist_retire', 'cat'] = 'capacity'

#Now pivot to turn each row into its own input
df_full = df_full.pivot_table(index=['jedi_scenario','tech','st', 'year','type'], columns='cat', values='Value').reset_index()

#add columns for outputs
df_full = df_full.reindex(columns=df_full.columns.values.tolist() + df_output_cat['output'].values.tolist())

#join output categories to outputs
df_outputs = pd.merge(left=df_outputs, right=df_output_cat, how='left', on=['output'], sort=False)

#loop through techs
for x, tech in enumerate(df_techs['tech'].values.tolist()):
    print('tech = ' + tech)
    #filter to just this tech
    df_tech = df_full[df_full['tech'] == tech]
    df_const = df_constants[df_constants['tech'] == tech]
    df_jedi_scen = df_jedi_scenarios[df_jedi_scenarios['tech'] == tech]
    df_var = df_variables[df_variables['tech']==tech]
    df_out = df_outputs[df_outputs['tech']==tech]
    df_out_oper = df_out[df_out['type']=='operation']

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
        df_exist = df_scen[df_scen['type']=='exist_retire']
        #set the inputs for existing plants
        for i, r in df_var.iterrows():
             ws_in.Range(r['cell']).Value = r['exist_val']
        #Grab the operation outputs for existing plants
        exist_out = {}
        for i,r in df_out_oper.iterrows():
            exist_out[r['output']] = float(ws_out.Range(r['cell']).Value)
        #Now fill the dataframe with existing outputs, scaled by project size
        for i, r in df_exist.iterrows():
            mult = r['capacity']/project_size
            for out in exist_out:
                df_full.loc[i, out] = mult*exist_out[out]

        #Now do new capacity
        df_new = df_scen[df_scen['type']=='new']
        #now, loop through df rows, fill in new capital and o&m cost, and get associated economic impacts
        c = 1
        for i, r in df_new.iterrows():
            costs = {}
            costs['cap_cost'] = r['cost_capital']/r['capacity']/1000
            costs['om_cost'] = r['cost_om']/r['capacity']/1000
            #set region as state if state_switch is True (WARNING: MAKE SURE ALL DATA IS MAPPED TO VALID STATES). Otherwise, United States will be used
            if state_switch:
                ws_in.Range('B13').Value = r['st'] #THIS NEEDS TO BE UPDATED TO USE CONSTANTS
            if c%10 == 0:
                print(str(c) + '/'+str(len(df_new)))
            #Set variable inputs
            for j, ro in df_var.iterrows():
                 ws_in.Range(ro['cell']).Value = costs[ro['cat']]
            #grab the economic outputs and scale by project size
            mult = r['capacity']/project_size
            for j,ro in df_out.iterrows():
                df_full.loc[i, ro['output']] = mult*float(ws_out.Range(ro['cell']).Value)
            c = c + 1
    wb.Close(False)

df_full.to_csv(this_dir + r'\outputs\df_full_pre_spread.csv', index=False)

#Add dataframes to map solve years to construction years and operation years.
solveyears = list(range(2010, 2052, 2))
oper_years = {'type': [], 'tech': [], 'year': [], 'econ_year': []}
constr_years = {'year':[], 'econ_year':[]}
for y in solveyears:
    constr_years['year'] += [y,y]
    constr_years['econ_year'] += [y-1,y]
    for i, r in df_techs.iterrows():
        for j in range(2010, y-1):
            oper_years['type'].append('exist_retire')
            oper_years['tech'].append(r['tech'])
            oper_years['year'].append(y)
            oper_years['econ_year'].append(j)
        for j in range(y, min(y + r['lifetime'], 2051)):
            oper_years['type'].append('new')
            oper_years['tech'].append(r['tech'])
            oper_years['year'].append(y)
            oper_years['econ_year'].append(j)
df_oper_years = pd.DataFrame(oper_years)
df_constr_years = pd.DataFrame(constr_years)

#Gather construction and operation outputs into separate lists
df_output_constr = df_output_cat[df_output_cat['type']=='construction']
df_output_oper = df_output_cat[df_output_cat['type']=='operation']
constr_outs = df_output_constr['output'].values.tolist()
oper_outs = df_output_oper['output'].values.tolist()

#Spread construction outputs over construction years, solve year t and t-1
df_constr_cols = [i for i in df_full.columns.values.tolist() if i not in oper_outs]
df_constr = df_full[df_constr_cols].copy()
df_constr['out_type'] = 'construction'
df_constr = df_constr[df_constr['type'] == 'new']
constr_renames = {}
for i,r in df_output_constr.iterrows():
    constr_renames[r['output']] = r['subtype']+'_'+r['subsubtype']
df_constr.rename(columns=constr_renames, inplace=True)
df_constr = pd.merge(left=df_constr, right=df_constr_years, how='left', on=['year'], sort=False)
#scale results by 1/2 to spread over two years
scale_cols = constr_renames.values() + ['capacity', 'cost_capital', 'cost_om']
df_constr.loc[:,scale_cols] = df_constr.loc[:,scale_cols]/2

#Spread operation outputs over operation years, solve year t to t+lifetime
oper_cols = [i for i in df_full.columns.values.tolist() if i not in constr_outs]
df_oper = df_full[oper_cols].copy()
df_oper['out_type'] = 'operation'
oper_renames = {}
for i,r in df_output_oper.iterrows():
    oper_renames[r['output']] = r['subtype']+'_'+r['subsubtype']
df_oper.rename(columns=oper_renames, inplace=True)
df_oper = pd.merge(left=df_oper, right=df_oper_years, how='left', on=['type', 'tech', 'year'], sort=False)
df_full = pd.concat([df_constr, df_oper], ignore_index=True)
#sort and rearrange columns
sortcols = ['jedi_scenario', 'tech', 'st', 'year', 'type', 'out_type', 'econ_year']
othercols = [i for i in df_full.columns.values.tolist() if i not in sortcols]
df_full = df_full[sortcols+othercols]
df_full = df_full.sort_values(sortcols)
df_full.to_csv(this_dir + r'\outputs\df_full_spread.csv', index=False)

#now sum across solve years a
df_full = df_full.groupby(['jedi_scenario', 'tech', 'st', 'type', 'out_type', 'econ_year'], as_index=False).sum()
outcols = [i for i in df_full.columns.values.tolist() if i not in ['capacity', 'cost_capital', 'cost_om', 'year']]
df_full = df_full[outcols]
df_full.to_csv(this_dir + r'\outputs\df_full_final.csv', index=False)