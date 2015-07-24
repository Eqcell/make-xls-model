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
    return read_sheet(filename_, sheet_, None).values.tolist()[0]  
    
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

def make_xl_model(abs_filepath, model_sheet): 
   print("Done importing libraries")
   
   data_df, controls_df, equations_list, var_list = get_spec_as_tuple(abs_filepath)
   print("Done reading specification from file")  
   for spec_element in get_spec_as_tuple(abs_filepath):
      pprint(spec_element)
   
   ar = make_wb_array2(data_df, controls_df, equations_list, var_list)
   # shorter notation for the above
   # ar = make_wb_array2(*get_spec_as_tuple(abs_filepath))
   print("Done creating 'wb_array'")   
   pprint(ar)
   
   write_array_to_xl_using_xlwings(ar, abs_filepath, model_sheet)  
   print("Finished writing to file: " + abs_filepath)

# def get_df_before_equations(data_df, controls_df, var_list):
    # df = pd.concat([data_df, controls_df])    
    # try:
        # return df[var_list]
    # except KeyError:
        # var_list_not_in_df = [x for x in var_list if x not in df.columns.values]
        # pprint (var_list_not_in_df)
        # raise KeyError("'var_list contains variables outside union of 'data_df' and 'contol_df' variables. \nCannot perform df[var_list]")    

# def make_wb_array(data_df, controls_df, equations_list, var_list):
    # df = get_df_before_equations(data_df, controls_df, var_list)
    # from xl_fill import 
   
###########################################################################
## Sample call
###########################################################################
    
if __name__ == '__main__':
    import os
    abs_filepath = os.path.abspath('spec2.xls')
    sheet = 'model'
    data_df, controls_df, equations_list, var_list = get_spec_as_tuple(abs_filepath)
    from xl_fill import get_dataframe_before_equations
    df = get_dataframe_before_equations(data_df, controls_df, var_list)
    pprint(df)
    
    #make_xl_model(abs_filepath, sheet)    
    
