import sys
import os
import pandas as pd
import numpy as np

df = pd.read_excel("sample_spec.xls", sheetname=0, header = None)
print (df)

ref_data_df = df.ix[1:4,1:3]
print(ref_data_df) 

ref_eq = df.ix[7:8, 1]

ref_controls = None


def get_sample_specification():
    model_spec = [
    ("Historic data as df",       ref_data_df ),
    ("Names as dict",             None        ),
    ("Equations as list",         ref_eq      ),
    ("Control parameters as df",  get_sample_controls_as_dataframe() )] 
    
    # requires workaround
    view_spec = [
    ['Excel filename' ,    'model.xls'],
    ['Sheet name' ,        'model'],
    ['List of variables',  ROW_LABELS_IN_OUTPUT] 
    ]




