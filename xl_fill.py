import numpy as np
import math
import pandas as pd
import re
from pprint import pprint

from data_source import get_sample_historic_data_as_dataframe, get_names_as_dict
from data_source import get_equations
from data_source import get_sample_controls_as_dataframe
from data_source import get_row_labels, get_years_as_list, get_xl_filename

TIME_INDEX_VARIABLES = ['t', 'T', 'n', 'N']

def get_dataframe_before_equations():    
    """
       This is a dataframe obtained by merging historic data and future values of control variables.
       WARNING: currently returns a stub. 
    """
    return pd.DataFrame(
          {   "GDP" : [66190.11992, 71406.3992, None, None]
          , "GDP_IQ": [101.3407976, 100.6404858, 95.0, 102.5]       
          , "GDP_IP": [105.0467483, 107.1941886, 115.0, 113.0]} 
          ,   index = [2013, 2014, 2015, 2016]
          )
          
def get_array_before_equations(df):
    """
       Decorate Dataframe with extra row (years) and column (var names) 
       and return as ndarray with object types. In resulting array some values for years
       will be NaN/nan. These are cells where Excel formulas need to be inserted.      
      
       Note: array of this kind directly represents an Excel worksheet.
             the intent is too fill this array and write it to Excel worksheet.
       
       Not todo: decorate also with a column of variable text descriptions (first column)
       """    
    ar0 = df.as_matrix().transpose().astype(object)
    labels = get_row_labels()
    ar = np.insert(ar0, 0, labels, axis = 1)
    years = [""] + get_years_as_list()
    ar = np.insert(ar, 0, years, axis = 0)
    return ar

def get_sample_array_after_equations():
    return np.array(   
    [['', 2013, 2014, 2015, 2016]
    ,['GDP', 66190.11992, 71406.3992, '=C2*D3*D4/10000', '=D2*E3*E4/10000']
    ,['GDP_IP', 105.0467483, 107.1941886, 115.0, 113.0]
    ,['GDP_IQ', 101.3407976, 100.6404858, 95.0, 102.5]]
    , dtype=object)
    
    # WARNING: actual intension was '=C2*D3/100*D4/100', '=C2*D3/100*D4/100'

   
def get_excel_formula(cell, equation):
    return "=B2"

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
        #eq = eq.strip()
        dependent_var, formula = eq.split('=')
        key = strip_timeindex(dependent_var)
        parsed_eq_dict[key] = {'dependent_var': dependent_var.strip(), 'formula': formula.strip()}
    return parsed_eq_dict

if __name__ == "__main__":
    import doctest
    doctest.testmod()
    
    # this is import of user-inputed variables from a different module
    data = get_sample_historic_data_as_dataframe()  
    names = get_names_as_dict()
    equations = get_equations()
    controls = get_sample_controls_as_dataframe()
    print(controls)
    row_labels = get_row_labels()
    years = get_years_as_list()
    
    # task formulation - inputs
    print("\n*** In this module on input we have:")
    
    print_dict = [ ["------ Data:",data],
                   ["------ Names:",names],
                   ["------ Equations:",equations], 
                   ["------ Control forecast parameters:",controls],
                   ["------ Years:",years],
                   ["------ Row labels:",row_labels]
                   ]
    for item in print_dict:
        print()
        print(item[0])
        pprint(item[1])
    
    # task formulation - final result
    print("\n*** Task - produce array with formulas (intent - it \"Excel-dumpable\"):")
    print(get_sample_array_after_equations()) 
    # end task formulation   
    
    # Solution:
    print("\n*** Solution flow:")
    df = get_dataframe_before_equations()
    ar = get_array_before_equations(df)
    
    print("\n*** Array before equations:")
    print(ar)
    
    print("\n***  Split formulas:")
    formulas = parse_to_formula_dict(equations)
    pprint(formulas)
    
    print("\n***  Assign variables to array rows:")
    
    # variables = parse_to_variables_dict(...)
    def get_variable_rows_in_array_as_dict(array, column = 0):
        variable_to_row_dict = {}
        for i, label in enumerate(array[:,column]):            
            variable_to_row_dict[label] = i            
        return variable_to_row_dict
    
    variable_to_row_dict = get_variable_rows_in_array_as_dict(ar)
    pprint(variable_to_row_dict)
    
    print("\n***  Formula creation:")
    def get_xl_formula(cell, var_name):
        """
        cell is (row, col) tuple
        varname is like 'GDP'
        """        
        equation = formulas[var_name]
        variables = get_variable_rows_in_array_as_dict(ar)
        time_period = cell[1]
        return parse_equation_to_xl_formula(equation, variables, time_period)
    
    #def parse_equation_to_xl_formula(equation, variables, time_period):
    #    return "=<xl-style formula here>"   
    
    from eqcell_core import parse_equation_to_xl_formula
    # parse_equation_to_xl_formula(dict_formula, dict_variables, column)
    
    print("\n***  Iterate over NaN in data area + fill with stub formula:")    
    
    def yield_cells_for_filling(ar):
        """
        TODO: must yeild coordinates of nan values from data area in *ar* 
              data area is all of ar, but not row 0 or col 0
        """
        row_offset = 1
        col_offset = 1
        
        # We loop and check which indexes correspond to nan
        for i, row in enumerate(ar[col_offset:]):
            for j, col in enumerate(row[row_offset:]):
                if math.isnan(col):
                    yield i + col_offset, j + row_offset 
    
    
    def get_var_label(ar, row, var_column = 0):
        return ar[row, var_column]
        # better - check is it is a valid variable name
        # var_list = unique(controls.columns.values.tolist() + row_labels)
    
    def fill_array_with_excel_equations(ar):
        for cell in yield_cells_for_filling(ar):
            var_name = get_var_label(ar, cell[0])
            ar[cell] = get_xl_formula(cell, var_name)
        return ar        

    ar = fill_array_with_excel_equations(ar)
    print("\n*** Current array:")
    print(ar)
         
    print("\n*** Must be like:")
    print(get_sample_array_after_equations())    

    is_equal = np.array_equal(ar, get_sample_array_after_equations())         
    
    print("\n*** Solution complete: " + str(is_equal))
    print()
    print(np.equal(ar, get_sample_array_after_equations())) 