import sys
import os
import pandas as pd
import numpy as np
from pprint import pprint

def read_specification_from_xls_file(filename):
    spec_dict = {} 
    for sheet_info in [['data',      0],
                       ['controls',  0],
                       ['equations', None],
                       ['format',    None]]:
        df = pd.read_excel(filename, sheetname=sheet_info[0], header =  sheet_info[1])
        spec_dict[sheet_info[0]] = df
    return spec_dict 
    
def vertical_1col_df_to_list(df):
    return df[0].values.tolist()
    
def get_specification(filename):
    spec_dict = read_specification_from_xls_file(filename)
    
    model_spec = [
    ("Historic data as df",       spec_dict['data'].transpose()),
    ("Names as dict",             None        ),
    ("Equations as list",         vertical_1col_df_to_list(spec_dict['equations']) ),
    ("Control parameters as df",  spec_dict['controls'].transpose())] 
    
    view_spec = [
    ['Excel filename' ,    'model.xls'],
    ['Sheet name' ,        'model'],
    ['List of variables',   vertical_1col_df_to_list(spec_dict['format'])]
    ]
    pprint(model_spec)     
    pprint(view_spec)
    
    return model_spec, view_spec 
    
if __name__ == '__main__':
    m, v = get_specification("spec.xls")
    print(m)
    print(v)