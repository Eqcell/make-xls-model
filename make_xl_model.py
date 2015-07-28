import pandas as pd
import numpy as np
from pprint import pprint
from xlwings import Workbook, Range, Sheet

from equations_preparser import parse_to_formula_dict
from iterate_in_array import fill_array_with_excel_formulas   
   
###########################################################################
## Import from Excel workbook (using pandas)
###########################################################################

def read_sheet(filename_, sheet_, header_):    
    return pd.read_excel(filename_, sheetname=sheet_, header = header_).transpose()
    
def read_df(filename_, sheet_):    
    return read_sheet(filename_, sheet_, 0)
    
def read_col(filename_, sheet_):    
    return read_sheet(filename_, sheet_, None).values.tolist()[0]  
    
def get_equations_as_dict(file):
    eq_list = read_col(file, 'equations')
    # todo: 
    #     parse_to_formula_dict must:
    #        - control left side of equations
    #        - supress comments
    #        - disregard lines without '='     
    return  parse_to_formula_dict(eq_list)
    
def get_spec_as_dict(file):   
    return   { 'data': read_df(file, 'data')    
       ,   'controls': read_df(file, 'controls') 
       ,  'equations': get_equations_as_dict(file)
       }
       
def get_spec_as_tuple(file): 
    s = get_spec_as_dict(file)
    return s['data'], s['controls'], s['equations']
 
###########################################################################
## Export to Excel workbook (using xlwings/pywin32)
###########################################################################

def write_array_to_xl_using_xlwings(ar, file, sheet):  
    # Note: if file is opened In Excel, it must be first saved before writing new output to it, but it may be left open in Excel application. 
    wb = Workbook(file)
    Sheet(sheet).activate()        
    Range(sheet, 'A1').value = ar.astype(str)    
    wb.save()

###########################################################################
## Grouped variables
###########################################################################
   
#def print_variable_names_by_group(group_dict):    
#    print ("Data vars:", group_dict['data'])
#    print ("Control vars:", group_dict['control'])
#    print ("Equation-derived vars:", group_dict['eq'])    
    
def get_variable_names_by_group(data_df, controls_df, equations_dict):
    """
    Obtain non-overlapping variable labels grouped into data, control 
    and equation-derived variables.    
    """
    
    # all variables from controls_df must persist in var_list (group 1 of variables)
    g1 = controls_df.columns.values.tolist()
    
    # group 2: variables in data_df not listed in control_df
    dvars = data_df.columns.values.tolist()
    g2 = [d for d in dvars if d not in g1]

    # group 3: variables on leftside of equations not listed in group 1 and 2
    evars = equations_dict.keys()
    g3 = [e for e in evars if e not in g1 + g2]
    
    return {'control': g1, 'data': g2, 'eq': g3}

###########################################################################
## Dataframe and array manipulation
###########################################################################


def make_empty_df(index_, columns_):
    df = pd.DataFrame(index=index_, columns=columns_)
    return  df 
    # Note can use: .fillna(0)

def subset_df(df, var_list):
    try:
        return df[var_list]
    except KeyError:        
        pprint ([x for x in var_list if x not in df.columns.values])
        raise KeyError("*var_list* contains variables outside *df* column names." +  
                       "\nCannot perform subsetting like df[var_list]")
    except:
        print ("Error handling dataframe:", df)
        raise ValueError
      
def make_df_before_equations(data_df, controls_df, equations_dict):
    """
    Return a dataframe containing all variable values from data, controls and 
    placeholder for new varaibales from equations.
    """    
    
    var_group = get_variable_names_by_group(data_df, controls_df, equations_dict)
    
    # concat data and control *df*
    df1 = data_df
    df2 = controls_df
    df = pd.concat([df1, df2])    
    
    # *df2* is a placeholder for equation-derived variables 
    df3 = make_empty_df(data_df.index, var_group['eq'])
    # add *df3* to *df*. 
    df = pd.merge(df, df3, left_index = True, right_index = True, how = 'left')
    
    # reorganise rows
    var_list = var_group['data'] +  var_group['eq'] + var_group['control']
    return subset_df(df, var_list)

def make_array_before_equations(df):
    """
    Convert dataframe to array, decorate with extra top row an extra left-side columns.
    Returns array and pivot column number. Pivot column contains variable labels.
    """
    ar = df.as_matrix().transpose().astype(object)
    
    # add variable labels as a first column in *ar*
    labels = df.columns.tolist()
    ar = np.insert(ar, 0, labels, axis = 1)
    pivot_col = 0

    # add years as first row in *ar*
    years = [""] + df.index.astype(str).tolist()
    ar = np.insert(ar, 0, years, axis = 0)
    
    return ar, pivot_col
    
###########################################################################
## Main entry point
###########################################################################

def get_input_variables(abs_filepath):
    data_df, controls_df, equations_dict = get_spec_as_tuple(abs_filepath) 
    var_group = get_variable_names_by_group(data_df, controls_df, equations_dict)
    return data_df, controls_df, equations_dict, var_group

def get_resulting_workbook_array(abs_filepath):
    
    # Get model specification 
    data_df, controls_df, equations_dict, var_group = get_input_variables(abs_filepath)    
    
    # Get array before formulas
    df = make_df_before_equations(data_df, controls_df, equations_dict)
    ar, pivot_col = make_array_before_equations(df)    
       
    # Fill array with formulas
    # Todo: fillable_var_list is effectively everything that appears on the left side of equations
    #       must compare *equations_dict* and *fillable_var_list*    
    fillable_var_list = var_group['data'] + var_group['eq']
    ar = fill_array_with_excel_formulas(ar, equations_dict, fillable_var_list, pivot_col)
    return ar

def make_xl_model(abs_filepath, sheet): 

    ar = get_resulting_workbook_array(abs_filepath)    
    print("Array to write to Excel sheet:")     
    print(ar) 
    write_array_to_xl_using_xlwings(ar, abs_filepath, sheet)
    
if __name__ == '__main__':
    
    import os
    for fn in ['spec.xls', 'spec2.xls']:
       abs_filepath = os.path.abspath(fn)
       sheet = 'model'
       make_xl_model(abs_filepath, sheet)
