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
state_vals_switch = True

this_dir = os.path.dirname(os.path.realpath(__file__))

#get reeds output data
dfs = gdxpds.to_dataframes(r"\\nrelqnap01d\ReEDS\FY17-JEDI-MRM-jedi\runs\JEDI 2017-08-02\gdxfiles\JEDI.gdx")

#Read in workbook input csvs
df_techs = pd.read_csv(this_dir + r'\inputs\techs.csv')
df_hierarchy = pd.read_csv(this_dir + r'\inputs\hierarchy.csv')
df_constants = pd.read_csv(this_dir + r'\inputs\constants.csv')
df_jedi_scenarios = pd.read_csv(this_dir + r'\inputs\jedi_scenarios.csv')
df_variables = pd.read_csv(this_dir + r'\inputs\variables.csv')
df_outputs = pd.read_csv(this_dir + r'\inputs\outputs.csv')
df_output_cat = pd.read_csv(this_dir + r'\inputs\output_categories.csv')
df_state_vals = pd.read_csv(this_dir + r'\inputs\state_vals.csv')

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

#Now pivot to turn each row into its own input
df_full = df_full.pivot_table(index=['jedi_scenario','tech','st', 'year'], columns='cat', values='Value').reset_index()

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
    df_st_vals = df_state_vals[df_state_vals['tech']==tech]

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

    #Gather wages and other state-level values
    reg_cell = df_techs['reg_cell'][x]
    st_vals = {}
    for st in df_tech['st'].unique():
        st_vals[st] = {}
        ws_in.Range(reg_cell).Value = st
        for j, ro in df_st_vals.iterrows():
            st_vals[st][ro['desc']] = ws_in.Range(ro['cell']).Value
    #reset region to united states
    ws_in.Range(reg_cell).Value = 'UNITED STATES'
    
    #loop through jedi scenarios, set inputs, and gather economic outputs
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
            if c%100 == 0:
                print(str(c) + '/'+str(len(df_scen)))
            #If we are using the state switch, simply select the state
            if state_switch:
                ws_in.Range(reg_cell).Value = r['st']
            #If not, we need to manually enter state-level values
            elif state_vals_switch == True:
                for j, ro in df_st_vals.iterrows():
                    ws_in.Range(ro['cell']).Value = st_vals[r['st']][ro['desc']]
            #Construction
            if pd.notnull(r['capacity_new']):
                #calculate input variables
                constr_vars = {}
                constr_vars['cap_cost'] = r['cost_capital']/r['capacity_new']/1000
                #set inputs in workbook
                for j, ro in df_var[df_var['type'] == 'construction'].iterrows():
                    ws_in.Range(ro['cell']).Value = constr_vars[ro['cat']]
                #gather outputs
                mult = r['capacity_new']/project_size
                for j,ro in df_out[df_out['type'] == 'construction'].iterrows():
                    df_full.loc[i, ro['output']] = mult*float(ws_out.Range(ro['cell']).Value)
            #Operation
            if pd.notnull(r['capacity_cumulative']):
                #calculate input variables
                oper_vars = {}
                oper_vars['om_cost'] = r['cost_om']/r['capacity_cumulative']/1000
                #set inputs in workbook
                for j, ro in df_var[df_var['type'] == 'operation'].iterrows():
                    ws_in.Range(ro['cell']).Value = oper_vars[ro['cat']]
                #gather outputs
                mult = r['capacity_cumulative']/project_size
                for j,ro in df_out[df_out['type'] == 'operation'].iterrows():
                    df_full.loc[i, ro['output']] = mult*float(ws_out.Range(ro['cell']).Value)
            c = c + 1
    wb.Close(False)

df_full.to_csv(this_dir + r'\outputs\df_full.csv', index=False)
