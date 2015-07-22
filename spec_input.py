from xlwings import Workbook, Range, Sheet
from pprint import pprint
import pandas as pd
from xl_fill import make_wb_array2
from xlwings import Workbook, Range, Sheet
   
###########################################################################
## Import from Excel workbook
###########################################################################

def read_sheet(filename_, sheet_, header_):    
    return pd.read_excel(filename_, sheetname=sheet_, header = header_).transpose()
    
def read_df(filename_, sheet_):    
    return read_sheet(filename_, sheet_, 0)
    
def read_col(filename_, sheet_):    
    return read_sheet(file, sheet_, None).values.tolist()[0]  
    
def get_spec_as_dict(file):
    return   { 'data': read_df(file, 'data')    
       ,   'controls': read_df(file, 'controls') 
       ,  'equations': read_col(file, 'equations') 
       ,     'format': read_col(file, 'format')   
       }
       
def get_spec_as_tuple(file): 
    s = get_spec_as_dict(file)
    return s['data'], s['controls'], s['equations'], s['format']
 
###########################################################################
## Export to Excel workbook
###########################################################################

def filter_(a):
    PRECISION = 6
    try:
        z = float(a)
        if round(z) == z:
           return int(z)
        else:
           return round(z, PRECISION)
    except ValueError:
        return a

def iterate_over_array(ar):
    for i, row in enumerate(ar):       
         for j, val in enumerate(row):
                yield i, j, val

def write_array_to_xl_using_xlwings(ar, file, sheet):    
    wb = Workbook(file)
    Sheet(sheet).activate()        
    # LATER: check why below does not work 
    # Range(sheet, 'A1').value = ar    
    for i, j, val in iterate_over_array(ar):
        Range(sheet, (i, j)).value = val  
    wb.save() 
    
###########################################################################
    
file = 'D:/make-xls-model-master/spec.xls'
data_df, controls_df, equations_list, var_label_list = get_spec_as_tuple(file)
wb_array = make_wb_array2(*get_spec_as_tuple(file))
sheet = 'model'

from data_source import _sample_for_xfill_array_after_equations as ar_sample
target_ar = ar_sample()

try: 
   write_array_to_xl_using_xlwings(target_ar, file, sheet)
except:
   print("ERROR")

try: 
   pass
   # write_array_to_xl_using_xlwings(wb_array, file, sheet)
except:
   print("ERROR")
   


###########################################################################