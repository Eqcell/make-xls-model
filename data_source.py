# coding: utf-8
"""
Data sources for the model and output.

Current entry point: 
    model_spec, view_spec = get_sample_specification()
        
Needed: 
    get_specification(model_user_param_dict, view_user_param_dict)
    
"""

from pprint import pprint
import pandas as pd

###########################################################################
## Proxy data (returned by get_sample* functions ) 
###########################################################################

# label, year, value
DATA_PROXY = [ ("GDP", 2013, 66190.11992)
        , ("GDP",    2014, 71406.3992)
        , ("GDP_IQ", 2013, 101.3407976)
        , ("GDP_IQ", 2014, 100.6404858)
        , ("GDP_IP", 2013, 105.0467483)
        , ("GDP_IP", 2014, 107.1941886) ] 

# label, year, value
CONTROLS_PROXY = [("GDP_IQ", 2015, 95.0)
        , ("GDP_IP", 2015, 115.0)
        , ("GDP_IQ", 2016, 102.5)
        , ("GDP_IP", 2016, 113.0)
        , ("is_forecast", 2015, 1)
        , ("is_forecast", 2016, 1)
        ]        
        
# title, label, group, level, precision
# ERROR: wont print cyrillic charactes, only whitespace.
NAMES_CSV_PROXY = [("ВВП",                      "GDP",    "Нацсчета", 1, 0),
                   ("Индекс физ.объема ВВП",    "GDP_IQ", "Нацсчета", 2, 1),
                   ("Дефлятор ВВП",	            "GDP_IP", "Нацсчета", 2, 1)]
 
EQ_SAMPLE = ["GDP(t) = GDP(t-1) * GDP_IP(t) / 100 * GDP_IQ(t) / 100"]

# change in test setting: one variable not in output 
ROW_LABELS_IN_OUTPUT = ["GDP", "GDP_IP", "GDP_IQ"] # , "is_forecast"]

# final_dataframe_proxy = """		2014	2015	2016
# ВВП	gdp	71406	=D3*E4/100*E5/100	=E3*F4/100*F5/100
# Индекс физ.объема ВВП	gdp_Iq	100,6	95	102
# Дефлятор ВВП	gdp_Ip	107,2	110	112
# """  

def get_sample_specification():
    model_spec = [
    ("Historic data as df",       get_sample_historic_data_as_dataframe() ),
    ("Names as dict",             get_sample_names_as_dict() ),
    ("Equations as list",         get_sample_equations() ),
    ("Control parameters as df",  get_sample_controls_as_dataframe() )] 
    
    # requires workaround
    view_spec = [
    ['Excel filename' ,    'model.xls'],
    ['Sheet name' ,        'model'],
    ['List of variables',  ROW_LABELS_IN_OUTPUT] 
    ]
    
    return model_spec, view_spec



###########################################################################
## General handling
###########################################################################
        
def make_dataframe_based_on_list_of_tuples(lt):
    """Returns a dataframe with years in rows and variables in columns. 
       *lt* is a list of tuples like *data_proxy* and *controls_proxy*"""  
    
    # Read dataframe
    df = pd.DataFrame(lt, columns=['prop', 'time', 'val'])
    # Pivot by time
    return df.pivot(index='time', columns='prop', values='val')
        
###########################################################################
## Historic data 
###########################################################################
        
def check_get_historic_data_as_dataframe():
    """
    >>> check_get_historic_data_as_dataframe()
    True
    """
    df1 = _internal_sample_historic_data_as_dataframe()
    df2 = get_sample_historic_data_as_dataframe()    
    # The following returns a dataframe object
    # return get_sample_historic_data_as_dataframe() == get_historic_data_as_dataframe()
    return df1.equals(df2)
       
def _internal_sample_historic_data_as_dataframe():
    # Used for testing in check_get_historic_data_as_dataframe
    z = { "GDP" : [66190.11992, 71406.3992 ]
          , "GDP_IQ": [101.3407976, 100.6404858]       
          , "GDP_IP": [105.0467483, 107.1941886]}
    return pd.DataFrame(z, index = [2013, 2014])

def get_sample_historic_data_as_dataframe():
    return make_dataframe_based_on_list_of_tuples(DATA_PROXY)
    
def get_historic_data_as_dataframe():
    pass

    
###########################################################################
## Names 
###########################################################################

def get_sample_names_as_dict():       
    return {x[1]:x[0] for x in NAMES_CSV_PROXY}

def get_names_as_dict():
    """Make name parameter dictionary callable by names_dict[label][param].
    """
    pass

###########################################################################
## Equations 
###########################################################################

def get_sample_equations():
    return EQ_SAMPLE

def get_equations():
    pass
  
###########################################################################
## Control parameters 
###########################################################################

def get_sample_controls_as_dataframe():
    return make_dataframe_based_on_list_of_tuples(CONTROLS_PROXY)
    
def get_controls_as_dataframe():
    pass
    
###########################################################################
## Years?
###########################################################################

def get_years_as_list():
    return [y for y in range(get_start_year(),get_max_control_year() + 1)]

def get_max_control_year():
    return max([y[1] for y in controls_proxy])

###########################################################################
## Output parameters - requires workaround
###########################################################################

# LIMITATION: One sheet per output Excel file
# def get_sheet_format():
    # return { 'filename':'macro.xls' 
        # , 'sheet': 'gdp_forecast'
        # , 'start_year': 2013
        # , 'rows': ["GDP", "GDP_IP", "GDP_IQ"]
        # }
        
# def get_xl_filename():
    # return get_sheet_format()['filename']
      
# def get_sheet_name():
    # return get_sheet_format()['sheet']
        
# def get_start_year():        
    # return get_sheet_format()['start_year']
        
# def get_row_labels():        
    # return get_sheet_format()['rows']   
      
def print_specification(model_and_view):
   for mv in model_and_view:               
        for spec in mv:
            print("\n------ {}:".format(spec[0]))
            pprint(spec[1])
      
if __name__ == "__main__":
    import doctest
    doctest.testmod()
    
    print_specification(get_sample_specification())             
