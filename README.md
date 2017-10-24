# ReEDS-JEDI link

## Purpose/Results
Running reeds_jedi.py produces JEDI economic outputs (jobs, earnings, value add, output) for a ReEDS scenario based on ReEDS output data in JEDI.gdx in the scenario's gdxfiles/ directory. The produced JEDI results are stored in a new gdx file, JEDI_out.gdx, which is dropped in the same gdxfiles/ directory. Multiple ReEDS scenarios may also be batched together.

## Dependencies/Setting up
- Windows
- GAMS 24.7
- A version of ReEDS that produces JEDI.gdx.
- Python 3.4 (until we update from GAMS 24.7). To get Python 3.4, download Anaconda for Python 3 at https://www.anaconda.com/download/
-- While downloading you can decide to add conda and python to your PATH. Otherwise, to use the commands, you'll have to navigate to the Anaconda3\Scripts\ directory. For example, mine is at "C:\Users\mmowers\AppData\Local\Continuum\Anaconda3\Scripts".
- Follow instructions at https://conda.io/docs/user-guide/tasks/manage-environments.html to create a conda environment for Python 3.4. In windows command prompt, this will look something like:
```bash
cd "C:\Users\mmowers\AppData\Local\Continuum\Anaconda3\Scripts"
conda create -n myenv34 python=3.4
activate myenv34
```
-- Note that I didn't have Anaconda3\Scripts\ added to my PATH, hence the navigation to \Anaconda3\Scripts in this code snippet.
- Now everytime you want to use myenv34 for installing packages, running python, and running .py scripts like reeds_jedi.py, you'll have to activate the environment. Again, if Anaconda3\Scripts\ is not added to your PATH, you'll have to navigate to Anaconda3\Scripts\ to activate the environment.
- With myenv34 activated:
-- install pandas and pywin32 packages with:
```bash
conda install pandas pywin32
```
-- Follow Elaine Hale's instructions at https://github.com/NREL/gdx-pandas#install to install GAMS python 3.4 bindings and gdx-pandas.

##File Structure
- reeds_jedi.py: This is the main script. It opens the jedi technology workbooks, iteratively enters ReEDS inputs from each JEDI.gdx, and collects the outputs in each JEDI_out.gdx. A set of switches at the top of this file allow for different behavior of the script (see comments for each switch).
- inputs/
-- reeds_scenarios.csv: Path(s) to ReEDS run directories or directories of ReEDS run directories
-- techs.csv: The JEDI technologies and associated JEDI workbooks in the jedi_models/ directory.
-- tech_map.csv: A mapping from ReEDS technologies ("bigQ") to JEDI techs ("tech").
-- constants.csv: For each JEDI technology, the cell values to be set as constant while iterating through JEDI scenarios and ReEDS results.
-- jedi_scenarios.csv: Different sets of content shares. An additional variable at the top of reeds_jedi.py allows these scenarios to be filtered for a given run.
-- variables.csv: For each JEDI technology, the cells associated with each ReEDS input.
-- state_vals.csv: By default, state-level wages are used in JEDI to reflect regional multipliers in ReEDS (even though national content shares are used). This file has the cells that are to be copied for a state.
-- outputs.csv: The set of outputs to be gathered for each iteration on ReEDS inputs.
-- output_categories.csv: This is a mapping from outputs in outputs.csv to their respective categories for type (construction, operation), outtype (jobs, earnings, output, value_add), and directtype (direct, indirect (aka supply chain), induced)
-- om_adjust.csv: For each JEDI technology, amounts that will be entered into the workbook and subtracted from operation and maintenance costs.
-- hierarchy.csv: A mapping from ReEDS PCAs to JEDI regions (US states + DC)

## Configuring and running
-- Modify reeds_scenarios.csv to point at the correct ReEDS run(s)
-- If needed, modify the content shares in jedi_scenarios.csv, and set the jedi_scenarios variable in the switches section of reeds_jedi.py
-- Modify any of the switches at the top of reeds_jedi.py
-- In a Windows command prompt, activate your python 3.4 environment (see above).
-- Navigate to the reeds_jedi folder and run reeds_jedi.py with
```bash
python reeds_jedi.py
```

## Exploring results
bokehpivot may be used to explore the data in JEDI_out.gdx, similar to how other ReEDS results are explored.

