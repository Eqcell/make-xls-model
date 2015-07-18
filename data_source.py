
# ***************************************************************************** 
# Proxies of data import form xls/csv/mysql (to speed up development / testing)
# *****************************************************************************

import io
import csv
from pprint import pprint
import pandas as pd

data_proxy = [ ("GDP", 2013, 66190.11992)
        , ("GDP",    2014, 71406.3992)
        , ("GDP_IQ", 2013, 101.3407976)
        , ("GDP_IQ", 2014, 100.6404858)
        , ("GDP_IP", 2013, 105.0467483)
        , ("GDP_IP", 2014, 107.1941886)
        ] 
        
def get_historic_data_as_dataframe():
    #todo: must return a dataframe based on *data_proxy*
    #z = { "GDP" : [66190.11992, 71406.3992 ]
    #      , "GDP_IQ": [101.3407976, 100.6404858]       
    #      , "GDP_IP": [105.0467483, 107.1941886]}
    # Read dataframe
    df = pd.DataFrame(data_proxy, columns=['prop', 'time', 'val'])
    # Pivot by time
    # pd.DataFrame(z, index = [2013, 2014])
    return df.pivot(index='prop', columns='time', values='val')
        
def get_data(label, year, data_proxy):
    slice = [x for x in data_proxy if x[0] == label]
    return  [x for x in slice if x[1] == year][0]


def yield_names_proxy():
    #names_csv_proxy ="""title	label	group	level	precision
    names_csv_proxy ="""ВВП	GDP	Нацсчета	1	0
Индекс физ.объема ВВП	GDP_IQ	Нацсчета	2	1
Дефлятор ВВП	GDP_IP	Нацсчета	2	1"""    
    text_stream = io.StringIO(names_csv_proxy)
    # better use csv.DictReader and work with dict in get_names_as_dict()
    reader = csv.reader(text_stream, delimiter='\t')
    for row in reader:
        yield row

def get_names_as_dict():
    """Make name parameter dictionary callable by names_dict[label][param].
    """
    names_dict = {}
    for row in yield_names_proxy():      
        sub_dict = {}
        sub_dict['title'] = row[0]
        sub_dict['precision'] = row[4]        
        new_entry_dict = {row[1]: sub_dict}
        names_dict.update(new_entry_dict)
    return names_dict

controls_proxy = [("GDP_IP", 2015, 95.0)
        , ("GDP_IQ", 2015, 115.0)
        , ("GDP_IP", 2016, 102.5)
        , ("GDP_IQ", 2016, 113.0)
        ]
        
def check_get_controls_as_dataframe():
    """
    >>> check_get_controls_as_dataframe() 
    True
    
    """
    return get_sample_controls_as_dataframe() == get_controls_as_dataframe()
    
def get_sample_controls_as_dataframe():    
    z = {'GDP_IP' : [95.0, 102.5],
         'GDP_IQ' : [115.0, 113.0]}   
    return pd.DataFrame(z, index=[2015, 2016])
    
def get_controls_as_dataframe():
    #todo: must return a dataframe based on *controls_proxy*
    #z = {'GDP_IP' : [95.0, 102.5],
    #     'GDP_IQ' : [115.0, 113.0]}   
    #return pd.DataFrame(z, index=[2015, 2016])
    # todo-again: must pass check_get_controls_as_dataframe(): doctest
    df = pd.DataFrame(controls_proxy, columns=['prop', 'time', 'val'])
    return df.pivot(index='prop', columns='time', values='val')


def get_years_as_list():
    return [y for y in range(get_start_year(),get_max_control_year() + 1)]

def get_max_control_year():
    return max([x[1] for x in controls_proxy])

def get_equations():
    return ["GDP(t) = GDP(t-1) * GDP_IP(t) / 100 * GDP_IQ(t) / 100"]

# LIMITATION: One sheet per output Excel file
def get_sheet_format():
    return { 'filename':'macro.xls' 
        , 'sheet': 'gdp_forecast'
        , 'start_year': 2013
        , 'rows': ["GDP", "GDP_IP", "GDP_IQ"]
        }
        
def get_xl_filename():
    return get_sheet_format()['filename']
      
def get_sheet_name():
    return get_sheet_format()['sheet']
        
def get_start_year():        
    return get_sheet_format()['start_year']
        
def get_row_labels():        
    return get_sheet_format()['rows']

    
final_dataframe_proxy = """		2014	2015	2016
ВВП	gdp	71406	=D3*E4/100*E5/100	=E3*F4/100*F5/100
Индекс физ.объема ВВП	gdp_Iq	100,6	95	102
Дефлятор ВВП	gdp_Ip	107,2	110	112
"""        

if __name__ == "__main__":
    import doctest
    doctest.testmod()

    data = get_historic_data_as_dataframe()  
    names = get_names_as_dict()
    equations = get_equations()
    controls = get_controls_as_dataframe()
    years = get_years_as_list()
    row_labels = get_row_labels()

    print_dict = { "------ Data:":data,
                   "------ Names:":names,
                   "------ Equations:": equations, 
                   "------ Control forecast parameters:":controls,
                   "------ Years:": years,
                   "------ Row labels:":row_labels
                   }
    for header, var in print_dict.items():
        print()
    #    print(header)
    #    pprint(var)

    #pprint(final_dataframe_proxy)
