import pandas as pd
from pprint import pprint
from xl_fill import make_wb_array2
from xlwings import Workbook, Range, Sheet
   
###########################################################################
## Import from Excel workbook (using pandas)
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
## Export to Excel workbook (using xlwings/pywin32)
###########################################################################

def write_array_to_xl_using_xlwings(ar, file, sheet):  
    # Note: must save file before opening   
    wb = Workbook(file)
    Sheet(sheet).activate()        
    Range(sheet, 'A1').value = ar.astype(str)    
    wb.save()
    
###########################################################################
## Main entry
###########################################################################
    
def make_xl_model(abs_filepath, sheet): 
    # data_df, controls_df, equations_list, var_label_list = get_spec_as_tuple(abs_filepath)
    ar = make_wb_array2(*get_spec_as_tuple(abs_filepath))
    write_array_to_xl_using_xlwings(ar, file, sheet)  

###########################################################################
## Sample call
###########################################################################
    
if __name__ == '__main__':
    # file = 'D:/make-xls-model-master/spec.xls'
    file = 'D:/git/make-xls-model/spec.xls'
    sheet = 'model'
    make_xl_model(file`, sheet)