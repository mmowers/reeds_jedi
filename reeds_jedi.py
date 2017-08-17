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
om_adjust_switch = True

jedi_scenarios = ['Low']

this_dir = os.path.dirname(os.path.realpath(__file__))

#get reeds output data
dfs = gdxpds.to_dataframes(r"\\nrelqnap01d\ReEDS\FY17-JEDI-MRM-jedi\runs\JEDI 2017-08-17b\gdxfiles\JEDI.gdx")

#Read in workbook input csvs
df_techs = pd.read_csv(this_dir + r'\inputs\techs.csv')
df_hierarchy = pd.read_csv(this_dir + r'\inputs\hierarchy.csv')
df_constants = pd.read_csv(this_dir + r'\inputs\constants.csv')
df_jedi_scenarios = pd.read_csv(this_dir + r'\inputs\jedi_scenarios.csv')
df_variables = pd.read_csv(this_dir + r'\inputs\variables.csv')
df_outputs = pd.read_csv(this_dir + r'\inputs\outputs.csv')
df_output_cat = pd.read_csv(this_dir + r'\inputs\output_categories.csv')
df_state_vals = pd.read_csv(this_dir + r'\inputs\state_vals.csv')
df_om_adjust = pd.read_csv(this_dir + r'\inputs\om_adjust.csv')

df_full = dfs['Jedi']
df_full.rename(columns={'jedi_tech': 'tech', 'allyears': 'year', 'jedi_cat': 'cat'}, inplace=True)

#convert text columns to lower case
df_full['tech'] = df_full['tech'].str.lower()
df_full['cat'] = df_full['cat'].str.lower()

#convert costs from 2004$ to 2015$
cost_cols = ['cost_capital', 'cost_om', 'cost_fuel', 'cost_var_om']
row_criteria = df_full['cat'].isin(cost_cols)
df_full.loc[row_criteria, 'Value'] = df_full.loc[row_criteria, 'Value'] / 0.796636801524834

#merge with states
df_full = pd.merge(left=df_full, right=df_hierarchy, how='left', on=['n'], sort=False)

#throw error if region is unmapped
df_unmapped = df_full[pd.isnull(df_full['st'])]
if not df_unmapped.empty:
    print(df_unmapped)
    sys.exit("unmapped regions shown above!")

#limit to only 2016 and after, and only US, and remove techs we aren't using
df_full = df_full[df_full['st'] != 'MEXICO']
df_full['year'] = df_full['year'].astype(int)
df_full = df_full[df_full['year'] >= 2016]
tech_list = df_techs['tech'].values.tolist()
df_full = df_full[df_full['tech'].isin(tech_list)]

#Test filter
if test_switch:
    df_full = df_full[df_full['st'] == 'ALABAMA']

#group and sum. Aggregate n to state in the process
df_full = df_full.groupby(['tech', 'st', 'year', 'cat'], as_index=False).sum()

#concatenate dataframe for each jedi scenario
df_full = df_full.reindex(columns=['jedi_scenario'] + df_full.columns.values.tolist())
df_temp = pd.DataFrame() #temporary
for scen_name in jedi_scenarios:
    df_full['jedi_scenario'] = scen_name
    df_temp = pd.concat([df_temp, df_full], ignore_index=True)
df_full = df_temp

#Now pivot to turn each row into its own input
index_cols = ['jedi_scenario','tech','st', 'year']
df_full = df_full.pivot_table(index=index_cols, columns='cat', values='Value').reset_index()

df_full.to_csv(this_dir + r'\outputs\df_in.csv', index=False)

#add columns for outputs
output_cols = df_output_cat['output'].values.tolist()
df_full = df_full.reindex(columns=df_full.columns.values.tolist() + output_cols)

#join output categories to outputs
df_outputs = pd.merge(left=df_outputs, right=df_output_cat, how='left', on=['output'], sort=False)

#loop through techs
for x, tech in enumerate(tech_list):
    print('tech = ' + tech)
    #filter to just this tech
    df_tech = df_full[df_full['tech'] == tech]
    df_const = df_constants[df_constants['tech'] == tech]
    df_jedi_scen = df_jedi_scenarios[df_jedi_scenarios['tech'] == tech]
    df_var = df_variables[df_variables['tech']==tech]
    df_out = df_outputs[df_outputs['tech']==tech]
    df_st_vals = df_state_vals[df_state_vals['tech']==tech]
    df_om_adj = df_om_adjust[df_om_adjust['tech']==tech]
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

    #Make adjustment to O&M based on property taxes, lease payments, etc.
    om_adjust = 0
    if om_adjust_switch == True:
        for i, r in df_om_adj.iterrows():
            ws_in.Range(r['cell']).Value = r['value']
            om_adjust += r['value']/(project_size*1000) #$/kW

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
                oper_vars['om_cost'] = r['cost_om']/r['capacity_cumulative']/1000 - om_adjust
                #for techs that have fuel cost, var om, etc:
                if pd.notnull(r['generation']):
                    oper_vars['var_om_cost'] = r['cost_var_om']/r['generation']
                    oper_vars['fuel_cost'] = r['cost_fuel']/r['fuel_use']
                    oper_vars['heat_rate'] = r['fuel_use']/r['generation']*1000
                    oper_vars['capacity_factor'] = r['generation']/(8760*r['capacity_cumulative'])
                #set inputs in workbook
                for j, ro in df_var[df_var['type'] == 'operation'].iterrows():
                    ws_in.Range(ro['cell']).Value = oper_vars[ro['cat']]
                #gather outputs
                mult = r['capacity_cumulative']/project_size
                for j,ro in df_out[df_out['type'] == 'operation'].iterrows():
                    df_full.loc[i, ro['output']] = mult*float(ws_out.Range(ro['cell']).Value)
            c = c + 1
    wb.Close(False)

#Remove inputs from output dataframe
df_full = df_full[index_cols + output_cols]
#Now adjust to account for non-solve years
min_year = df_full['year'].min()
max_year = df_full['year'].max()
#Reshape dataframe so that years are columns and output categories are rows
df_full = pd.melt(df_full, id_vars=index_cols, value_vars=output_cols, var_name='output', value_name= 'value')
index_cols.remove('year')
df_full = df_full.pivot_table(index=index_cols+['output'], columns='year', values='value').reset_index()
df_full.columns.name = None
df_full.fillna(0, inplace=True)
#merge with output categories
df_full = pd.merge(left=df_full, right=df_output_cat, how='left', on=['output'], sort=False)
#For construction, solve year t econimic outputs are halved, and the remaining half is assigned to non-solve year t-1
years = list(range(min_year, max_year+2, 2))
constr_rows = df_full['type'] == 'construction'
df_full.loc[constr_rows, years] = df_full.loc[constr_rows, years]/2
#now add the odd years and fill
oper_rows = df_full['type'] == 'operation'
for y in list(range(min_year+1, max_year+1, 2)):
    #default to construction, where non-solve year t is the same as solve year t+1
    df_full[y] = df_full[y+1]
    #For operation, non-solve year t econimic outputs are the average of solve years t-1 and t+1
    df_full.loc[oper_rows, y] = (df_full.loc[oper_rows, y-1] + df_full.loc[oper_rows, y+1])/2
df_full = pd.melt(df_full, id_vars=index_cols+df_output_cat.columns.values.tolist(), value_vars=list(range(min_year, max_year+1)), var_name='year', value_name= 'value')
#df_full = df_full.reindex(columns = index_cols+df_output_cat.columns.values.tolist()+list(range(min_year, max_year+1)))
df_full = df_full[pd.notnull(df_full['value'])]
df_full.to_csv(this_dir + r'\outputs\df_out.csv', index=False)
