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
   print("\n***** Step 1/4")
   print("Done importing libraries")
   
   data_df, controls_df, equations_list, var_list = get_spec_as_tuple(abs_filepath)
   print("\n***** Step 2/4")
   print("Done reading specification from file")  
   for spec_element in get_spec_as_tuple(abs_filepath):
      print()
      pprint(spec_element)
   
   ar = make_wb_array2(data_df, controls_df, equations_list, var_list)
   # shorter notation for the above
   # ar = make_wb_array2(*get_spec_as_tuple(abs_filepath))
   print("\n***** Step 3/4")
   print("Done creating 'wb_array'")   
   pprint(ar)
   
   # write_array_to_xl_using_xlwings(ar, abs_filepath, model_sheet)
   print("\n***** Step 4/4")
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
    data_df, controls_df, equations_list, user_var_list = get_spec_as_tuple(abs_filepath)

    # all variables from controls_df must persist in var_list (group 1 of variables)
    g1 = controls_df.columns.values.tolist()
    print (g1)
    print (controls_df[g1])
    # group 2: variables in data_df not listed in control_df
    dvars = data_df.columns.values.tolist()
    g2 = [d for d in dvars if d not in g1]
    print (g2)

    # group 3: variables on leftside of equations not listed in group 1 and 2
    from equations_preparser import parse_to_formula_dict
    evars = parse_to_formula_dict(equations_list).keys()
    g3 = [e for e in evars if e not in g1 + g2]
    print(g3)
    # make empty array based on g3 varnames and g2 dates
    df3 = pd.DataFrame(index=data_df.index, columns=g3)
    df3 = df3.fillna(0)
    print(df3)
    
    
    var_list = g2 + g1 + g3

    from xl_fill import get_dataframe_before_equations
    # df = get_dataframe_before_equations(data_df, controls_df, var_list)
    #pprint(df)
    
    # We first concatenate columns
    df2 = pd.concat([data_df[g2], controls_df[g1]])
    df = pd.merge(df2, df3, left_index = True, right_index = True, how = 'left')
    print("df: ",df[var_list])

    # Subsetting a union of 'data_df' and 'controls_df' and 'df3', protected for error.
    try:
        df = df[var_list]
    except KeyError:
        var_list_not_in_df = [x for x in var_list if x not in df.columns.values]
        pprint (var_list_not_in_df)
        raise KeyError("'var_list contains variables outside union of 'data_df' and 'contol_df' variables. \nCannot perform df[var_list]")
    except:
       print ("Error handling dataframes in get_dataframe_before_equations() in xl_fill.py")
       # LATER: add actual name of this file, obtained as a function
       # return None
    
    import numpy as np
    ar = df.as_matrix().transpose().astype(object)

    labels = df.columns.tolist()
    ar = np.insert(ar, 0, labels, axis = 1)
    pivot_col = 0
    years = [""] + df.index.astype(str).tolist()
    ar = np.insert(ar, 0, years, axis = 0)
   
    from new_iter import fill_array_with_excel_formulas
    ar = fill_array_with_excel_formulas(ar, equations_list, g2 + g3, var_list)
    print(ar)