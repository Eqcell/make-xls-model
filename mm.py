from data_source import get_specifcation, get_sample_specification
from xl_fill import make_wb_array
from backend import write_output_to_xls
from yaml_parser import get_user_param, get_default_param

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
model_dict, view_dict = get_sample_specifcation(None, None)


####### 
####### 3. Do main job
#######  
# create an numpy array, representing resulting worksheet
wb_array = make_wb_array(model_dict, view_dict)


####### 
####### Write results to file 
#######  
# write array to output excel file 
# note: upon implementation can be a new sheet in same file (e.g. xlwings, even for an open file)
write_output_to_xls(wb_array, view_dict)

#
#     if no user specification - write all controls in orginal order,
#                              - followed by data without controls in orginal order
#                              - year depth as in historic data
#