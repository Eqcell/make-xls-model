# coding: utf-8
"""
   Generate Excel file with ordered rows containing Excel formulas 
   that allow to calculate forecast values based on historic data, 
   equations and forecast parameters. Order of rows in Excel file 
   controlled by template definition. Start year specified as input.

   Input:  
        data
        equations
        names
        controls (forecast parameters)
        formats 
           xl_filename
           sheet
           start_year
           row_labels        
        
   Output: 
        macro.xls
        (an array of values to be written to macro.xls)
"""

import numpy as np
import pandas as pd
import re
from pprint import pprint

from data_source import get_sample_specification, print_specification
from eqcell_core import parse_equation_to_xl_formula, TIME_INDEX_VARIABLES

def check_get_dataframe_before_equations():
    """
    >>> check_get_dataframe_before_equations()
    True
    """
    df1 = _internal_get_dataframe_before_equations()    
    df2 = get_dataframe_before_equations()    
    return df1.equals(df2)


def get_dataframe_before_equations(data_df = None, controls_df = None, var_label_list = None):    
    """
       This is a dataframe obtained by merging historic data and future values of control variables.      
    """
    # TODO: must merge data_df, controls_df into a common dataframe
    #       years are extended to include both data_df years and controls_df years
    #       missing values are None
    #       order of columns is same as listed in var_label_list
    #
    #       not todo: resove possible conflicts in data_df/control_df columns and  var_label_list   
    #       not todo: default behaviour in column first lists data_df, then elements of controls_df, which are not in controls_df ('is_forcast' in example)
    #       not todo: no check of years continuity
    #pprint(data_df)
    #pprint(controls_df)
    #pprint(var_label_list)
    
    # We first concatenate columns
    df = pd.concat([data_df, controls_df])
    
    # Subsetting a union of 'data_df' and 'controls_df', protected for error.
    try: 
       return df[var_label_list]
    except ValueError:
       print ("Error handling dataframes in get_dataframe_before_equations()")
       return None
    

def _internal_get_dataframe_before_equations():    
    """
       This is a dataframe obtained by merging historic data and future values of control variables.
       WARNING: currently returns a stub. 
    """
    return pd.DataFrame(
          {   "GDP" : [66190.11992, 71406.3992, None, None]
          , "GDP_IP": [101.3407976, 100.6404858, 95.0, 102.5]       
          , "GDP_IQ": [105.0467483, 107.1941886, 115.0, 113.0]
          # Test setting: dataframe before equations has less columns than union of controls and data
          # , "is_forecast": [None, None, 1, 1] 
          } 
          ,   index = [2013, 2014, 2015, 2016]
          )
          
def get_array_before_equations(df):
    """
       Decorate *df* with extra row (years) and column (var names) 
       and return as ndarray with object types. In resulting array some values for years
       will be NaN/nan. These are cells where Excel formulas need to be inserted.      
      
       Note: array of this kind directly represents an Excel worksheet.
             the intent is too fill this array and write it to Excel worksheet.
       
       Not todo: decorate also with a column of variable text descriptions (first column)
       """    
    ar = df.as_matrix().transpose().astype(object)
    labels = df.columns.tolist() 
    ar = np.insert(ar, 0, labels, axis = 1)
    years = [""] + df.index.tolist()
    ar = np.insert(ar, 0, years, axis = 0)
    return ar

def get_sample_array_after_equations():
    return np.array(   
    [['', 2013, 2014, 2015, 2016]
    ,['GDP', 66190.11992, 71406.3992, '=C2*D3*D4/10000', '=D2*E3*E4/10000']
    ,['GDP_IP', 105.0467483, 107.1941886, 115.0, 113.0]
    ,['GDP_IQ', 101.3407976, 100.6404858, 95.0, 102.5]
    ,['is_forecast', "", "", 1, 1]]
    , dtype=object)
    
    # WARNING: actual intention was '=C2*D3/100*D4/100', '=C2*D3/100*D4/100'
   

def unique(list_):
    """Returns unique elements from list.
    >>> unique(['a','a'])
    ['a']
    """
    return list(set(list_))
 
def strip_timeindex(str_):
    """Returns variable name without time index.
    TODO: if function cannot strip time index anything return None

    Tests:
    >>> strip_timeindex("GDP(t)")
    'GDP'

    WARNING: in no time index return 'str_'. Test below fails.
    #>>> strip_timeindex("GDP")
    #'GDP'
    """
    all_indices = "".join(TIME_INDEX_VARIABLES)
    pattern = r"(\S*)[\[(][" + all_indices + "][)\]]"
    m = re.search(pattern, str_)
    if m:
        return m.groups()[0]
    else:
        return None
        
def test_parse_to_formula_dict():    
    """
    >>> test_parse_to_formula_dict()
    True
    True
    True
    """
    inputs = [
      ['GDP(t) = GDP(t-1) * GDP_IP(t) / 100 * GDP_IQ(t) / 100']
    , ['x(t) = x(t-1) + 1']
    , ['x(t) = x(t-1) + 1', 'y(t) = x(t)']
    ]    
    expected_outputs = [
      {'GDP': {'dependent_var': 'GDP(t)', 'formula': 'GDP(t-1) * GDP_IP(t) / 100 * GDP_IQ(t) / 100'}}   
    , {'x': {'dependent_var': 'x(t)', 'formula': 'x(t-1) + 1'}}
    , {'x': {'dependent_var': 'x(t)', 'formula': 'x(t-1) + 1'}, 'y': {'dependent_var': 'y(t)', 'formula': 'x(t)'}}
    ]
    for input_eq, expected_output in zip(inputs,expected_outputs):
       print(expected_output == parse_to_formula_dict(input_eq))

def parse_to_formula_dict(equations):
    """Returns a dict with left and right hand side of equations, referenced by variable name in keys."""
    parsed_eq_dict = {}
    for eq in equations:
        dependent_var, formula = eq.split('=')
        key = strip_timeindex(dependent_var)
        parsed_eq_dict[key] = {'dependent_var': dependent_var.strip(), 'formula': formula.strip()}
    return parsed_eq_dict

def get_variable_rows_as_dict(array, column = 0):
        variable_to_row_dict = {}        
        for i, label in enumerate(array[:,column]):           
            variable_to_row_dict[label] = i              
        #LATER: cut off one row (with years)
        #LATER: compare to full variable list
        return variable_to_row_dict

def get_xl_formula(cell, var_name):
        """
        cell is (row, col) tuple
        varname is like 'GDP'
        """ 
        try:        
            equation = formulas_dict[var_name]
            variables = get_variable_rows_as_dict(ar)
            time_period = cell[1]
            print (time_period, equation, variables)
            return parse_equation_to_xl_formula(equation, variables, time_period)
        except KeyError:
            return ""
    
def get_var_label(ar, row, var_column = 0):
        return ar[row, var_column]
        # better - check is it is a valid variable name
        # var_list = unique(controls.columns.values.tolist() + row_labels)

def fill_array_with_excel_equations(ar):
        for cell in yield_cells_for_filling(ar):
            var_name = get_var_label(ar, cell[0])
            ar[cell] = get_xl_formula(cell, var_name)
        return ar        

def get_xl_formula(cell, var_name, formulas_dict, variables_dict):
        """
        cell is (row, col) tuple
        varname is like 'GDP'
        """ 
        try:        
            equation = formulas_dict[var_name]
            time_period = cell[1]            
            return parse_equation_to_xl_formula(equation, variables_dict, time_period)
        except KeyError:
            return ""        

def get_var_label(ar, row, var_column = 0):
        return ar[row, var_column]
        # better - check is it is a valid variable name
        # var_list = unique(controls.columns.values.tolist() + row_labels)
    
def yield_cells_for_filling(ar):
    """
    Yields coordinates of nan values from data area in *ar* 
    Data area is all of ar, but not row 0 or col 0
          
    Example:
    
    >>> gen = yield_cells_for_filling([['', 2013, 2014, 2015, 2016],
    ...                                ['GDP', 66190.11992, 71406.3992, np.nan, np.nan],
    ...                                ['GDP_IP', 105.0467483, 107.1941886, 115.0, 113.0],
    ...                                ['GDP_IQ', 101.3407976, 100.6404858, 95.0, 102.5]])
    >>> next(gen)
    (1, 3)
    >>> next(gen)
    (1, 4)
    
    """
    row_offset = 1
    col_offset = 1
    
    # We loop and check which indexes correspond to nan
    for i, row in enumerate(ar[col_offset:]):
        for j, col in enumerate(row[row_offset:]):
            if np.isnan(col):
            # if math.isnan(col):
                yield i + col_offset, j + row_offset     
    
def fill_array_with_excel_formulas(ar, formulas_dict):
        variables_dict = get_variable_rows_as_dict(ar)
        for cell in yield_cells_for_filling(ar):
            var_name = get_var_label(ar, cell[0])
            ar[cell] = get_xl_formula(cell, var_name, formulas_dict, variables_dict)
        return ar  

def check_wb_array():
    """
    >>> check_wb_array()
    True
    """
    ar1 = make_wb_array()
    ar2 = get_sample_array_after_equations()    
    return np.array_equal(ar1, ar2)
        
def make_wb_array(model_dict = None, view_dict = None):
    # WARNING: must change to required arguements later
    if model_dict is None or view_dict is None:
        model_spec, view_spec = get_sample_specification()
        
    # WARNING: names_dict not used
    [data_df, names_dict, equations_list, controls_df] = [s[1] for s in model_spec]
    [xl_file, sheet, var_label_list] = [s[1] for s in view_spec]

    df = get_dataframe_before_equations(data_df, controls_df, var_label_list)
    ar = get_array_before_equations(df)

    formulas_dict = parse_to_formula_dict(equations_list)
    ar = fill_array_with_excel_formulas(ar, formulas_dict)        

    return ar
    
        
if __name__ == "__main__":

    # unpack variables locally 
    model_spec, view_spec = get_sample_specification()
    [data_df, names_dict, equations_list, controls_df] = [s[1] for s in model_spec]
    [xl_file, sheet, var_label_list] = [s[1] for s in view_spec]
    
    # task formulation - inputs
    print("\n****** Module inputs:")
    
    print_specification(get_sample_specification()) 
    
    # task formulation - final result
    print("\n****** Task - produce array with values and string-like formulas (intent - later write it to Excel)")
    print("Target array:")
    print(get_sample_array_after_equations()) 
    # end task formulation   
    
    # Solution:
    print("\n******* Solution flow:")    
    print("*** Array before equations:")
    df = get_dataframe_before_equations(data_df, controls_df, var_label_list)
    ar = get_array_before_equations(df)
    print(ar)
    
    print("\n***  Split formulas:")
    formulas_dict = parse_to_formula_dict(equations_list)
    pprint(formulas_dict)
    
    print("\n***  Assign variables to array rows:")
    variable_to_row_dict = get_variable_rows_as_dict(ar)
    pprint(variable_to_row_dict)
    
    print("\n***  Iterate over NaN in data area + fill with stub formula:")    
    ar = fill_array_with_excel_formulas(ar, formulas_dict)
    
    print("\n*** Resulting array:")
    print(ar)
         
    print("\n*** Must be like:")
    print(get_sample_array_after_equations())    

    is_equal = np.array_equal(ar, get_sample_array_after_equations())         
    
    print("\n*** Solution complete: " + str(is_equal))
    print()
    print(np.equal(ar, get_sample_array_after_equations()))
    
    import doctest
    doctest.testmod()
