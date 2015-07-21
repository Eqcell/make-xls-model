"""Make Excel model from specified historic data, equations and forecast control variables. 
Produces xls(x) <OUTPUT_XLS_FILE> file with data, controls and formulas inside it (no dependcies, no VBA).

Usage:   
   mxm.py --selftest
   mxm.py <SPECIFICATION_XLS_FILE> [<OUTPUT_XLS_FILE>]
   mxm.py --markup <YAML_FILE>
"""


"""
TO BE CORRECTED:
- no variable names written to file
- no running of all doctest tests  
- 

LIMITATIONS (by design):
- value precisions in output xls file not corrected (can be 0 and 2)
- one sheet ber file
- decorations of dt_before_equations are hard-coded
- eats cyrulic in NAMES
"""

"""
Notes:


   Generate Excel file with ordered rows containing Excel formulas 
   that allow to calculate forecast values based on historic data, 
   equations and forecast parameters. Order of rows in Excel file 
   controlled by template definition. Start year specified as input.

   Input:  
        data
        equations
        names
        controls (forecast parameters)
        formats 
           xl_filename
           sheet
           start_year
           row_labels        
           
   Entry point to module:
        
   Output: 
        macro.xls
        (an array of values to be written to macro.xls)

"""


from docopt import docopt
from data_source import get_mock_specification #, get_specification
from xl_fill import make_wb_array
from xls_io import write_output_to_xls
#from yaml_parser import get_user_param, get_default_param

arg = docopt(__doc__)
# print(arg)

if arg['--selftest']:
    model_spec, view_spec = get_mock_specification()
else:
    # temp script end
    raise ValueError("No inputs provided for script. Only 'mxm.py --selftest' option supported") 

# main job of creating resulting workbook array 'wb_array', dumpable to Excel
wb_array = make_wb_array(model_spec, view_spec)

# dump 'wb_array' to Excel
write_output_to_xls(wb_array, view_spec)

    


####### 
####### 1. Get user input
#######  
# get user input based on yaml config file
# yaml_filepath = "frame1_markup.txt"
# model_user_param_dict, view_user_param_dict = get_user_param(yaml_filepath)

# get user input based on two filenames only and default spec file formatting
# specification_xl_file = "spec.xls"
# output_xl_file = "model.xls"
# model_user_param_dict, view_user_param_dict = get_default_param(specification_xl_file, output_xl_file)

####### 
####### 2. Set model and output specification 
#######  
# make model and output specification based on user input
# model_dict, view_dict = get_specification(model_user_param_dict, view_user_param_dict)
# get defaults 
# model_spec, view_spec = get_sample_specification()
# print(model_spec)
# print(view_spec)


####### 
####### 3. Do main job
#######  
# create an numpy array, representing resulting worksheet
# wb_array = make_wb_array() # (wb_array, view_spec)


####### 
####### 4. Write results to file 
#######  
# write array to output excel file 
# note: upon implementation can be a new sheet in same file (e.g. xlwings, even for an open file)
#write_output_to_xls(wb_array, view_spec)

#
#     if no user specification - write all controls in orginal order,
#                              - followed by data without controls in orginal order
#                              - year depth as in historic data
#

#
# 2015-07-20 10:46 AM
#
# Outline:
# - write to xls
# - one input xls file
# - larger xls file (flatten folder)
# - experiment with actual data - full cycle (flatten folder)
# - (source + markup) vs flattened db vs db dump
#
