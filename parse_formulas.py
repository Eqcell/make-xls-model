from collections import deque
from xlwings import Workbook, Range, Sheet
from sympy import var
import os
# from docopt import docopt
import itertools

from data_source import get_historic_data_as_dataframe, get_names_as_dict
from data_source import get_equations
from data_source import get_controls_as_dataframe
from data_source import get_row_labels, get_years_as_list, get_xl_filename

import numpy as np
import pandas as pd

TIME_INDEX_VARIABLES = ['t', 'T', 'n', 'N']

data = get_historic_data_as_dataframe()  
names = get_names_as_dict()
equations = get_equations()
controls = get_controls_as_dataframe()
row_labels = get_row_labels()
years = get_years_as_list()

def unique(list_):
       return list(set(list_))
        
def get_dataframe_before_equations():    
    return pd.DataFrame(
          {   "GDP" : [66190.11992, 71406.3992, None, None]
          , "GDP_IQ": [101.3407976, 100.6404858, 95.0, 102.5]       
          , "GDP_IP": [105.0467483, 107.1941886, 115.0, 113.0]} 
          ,   index = [2013, 2014, 2015, 2016]
          )
def get_array_defore_equations():
    df = get_dataframe_before_equations() 
    ar0 = df.as_matrix().transpose().astype(object) #
    labels = get_row_labels()
    print(labels)
    ar = np.insert(ar0, 0, labels, axis = 1)
    years = [""] + get_years_as_list()
    ar = np.insert(ar, 0, years, axis = 0)
    return ar
    
         
df = get_dataframe_before_equations()
ar = get_array_defore_equations()
formulas = equations
variables = unique(controls.columns.values.tolist() + row_labels)

print(df)
print(ar)
print (formulas)
print (variables)

    

# ***************** 'eqcell' code below *****************
    

    # start_cell, end_cell = find_start_end_indices(wb, is_contiguous)
    # # Parse the column under START label
    # variables, formulas, comments = parse_variable_values(wb, start_cell, end_cell)
    # # checks if 'is_forecast' variable name is present on sheet
    # require_variable(variables, 'is_forecast')
    # # parse formulas and obtain Sympy expressions
    # parsed_formulas = parse_formulas(formulas, variables)

    # apply_formulas_on_sheet(wb, variables, parsed_formulas, start_cell)
    # if savefile is None:
        # savefile = workfile
    # save_workbook(wb, savefile)
    

def parse_formulas(formulas, variables):
    """
    Takes formulas as a dict of strings and returns a dict
    where dependent (left-hand side) variable and (right-hand side) formula
    are separated and converted to sympy expressions.
    input variable example:
    formulas = {'a(t)=a(t-1)*a_rate(t)': 6, 'b(t)=b_share(t)*a(t)': 11}
    output example:
    formulas_dict = {5: {'dependent_var': a(t), 'formula': a(t-1)*a_rate(t)},
                     9: {'dependent_var': b(t), 'formula': b(t-1)+2}}
    5, 6, 9, 11 are the row indices in the sheet. Row indices in formulas_dict changed to rows with variables.
    These rows contain data and are used to fill in formulas in forecast period.
    a(t), b(t-1)+2, ... are sympy expressions.
    """
    varirable_list = list(variables.keys()) + TIME_INDEX_VARIABLES

    # declares sympy variables
    var(' '.join(varirable_list))
    parsed_formulas = dict()

    for formula_string in formulas.keys():
        # removing white spaces
        formula_string = formula_string.strip()
        dependent_var, formula = formula_string.split('=')
        dependent_var = evaluate_variable(dependent_var)
        formula = evaluate_variable(formula)
        # finding the row where the formula will be applied - may be a function
        row_index = variables[str(dependent_var.func)]
        parsed_formulas[row_index] = {'dependent_var': dependent_var, 'formula': formula}

    return parsed_formulas
    
def parse_variable_values(workbook, start_cell, end_cell):
    """
    Given the workbook reference and START-END index pair, this function parses the values in the variable row
    and saves it as a list of the same name.
    input
    -----
    workbook:   Workbook xlwings object
    start_cell: Start cell dictionary
    end_cell:   End cell dictionary
    returns:    lists of variables, formulas, comments
    """
    workbook.set_current()    # sets the workbook as the current working workbook
    variables = dict()
    formulas = dict()
    comments = dict()
    start = (start_cell['row'], start_cell['col'])
    end = (end_cell['row'],   start_cell['col'])

    start_column = Range(get_sheet(), start, end).value
    # [1:] excludes 'START' element
    start_column = start_column[1:]

    for relative_index, element in enumerate(start_column):
        current_index = start_cell['row'] + relative_index + 1
        if element:    # if non-empty
            if not isinstance(element, str):
                raise ValueError("The column below START can contain only strings")

            # print(element)
            element = element.strip()

            if '=' in element:
                formulas[element] = current_index
            elif '#' == element[0]:
                comments[element] = current_index
            else:
                variables[element] = current_index

    return variables, formulas, comments

def require_variable(variables, var='is_forecast'):
    """
    Checks if variable string (default: `is_forcast`) is in the sheet variables dict, else raises error
    input
    -----
    variables: A dict of variables from excel sheet
    var:       A variable name string, to be checked if exists in variables.
    """
    if var not in variables.keys():
        raise ValueError('is_forecast is a mandatory value under START cell in excel sheet')

def evaluate_variable(x):
        try:
            x = eval(x)     # converting the formula into sympy expressions
        except NameError:
            raise NameError('Undefined variables in formulas, check excel sheet')
        return x

def parse_formulas(formulas, variables):
    """
    Takes formulas as a dict of strings and returns a dict
    where dependent (left-hand side) variable and (right-hand side) formula
    are separated and converted to sympy expressions.
    input variable example:
    formulas = {'a(t)=a(t-1)*a_rate(t)': 6, 'b(t)=b_share(t)*a(t)': 11}
    output example:
    formulas_dict = {5: {'dependent_var': a(t), 'formula': a(t-1)*a_rate(t)},
                     9: {'dependent_var': b(t), 'formula': b(t-1)+2}}
    5, 6, 9, 11 are the row indices in the sheet. Row indices in formulas_dict changed to rows with variables.
    These rows contain data and are used to fill in formulas in forecast period.
    a(t), b(t-1)+2, ... are sympy expressions.
    """
    varirable_list = list(variables.keys()) + TIME_INDEX_VARIABLES

    # declares sympy variables
    var(' '.join(varirable_list))
    parsed_formulas = dict()

    for formula_string in formulas.keys():
        # removing white spaces
        formula_string = formula_string.strip()
        dependent_var, formula = formula_string.split('=')
        dependent_var = evaluate_variable(dependent_var)
        formula = evaluate_variable(formula)
        # finding the row where the formula will be applied - may be a function
        row_index = variables[str(dependent_var.func)]
        parsed_formulas[row_index] = {'dependent_var': dependent_var, 'formula': formula}

    return parsed_formulas

def simplify_expression(expression, time_period, variables, depth=0):
    # get_variable_to_cell_segments
    """
    A recursive function which breaks a Sympy expression into segments, where each segment points to one cell on the
    excel sheet upon substitution of time index variable (t). Returns a dictionary of such segments and the computed
    cells.
    input
    -----
    expression:       Sympy expression, e.g: a(t - 1)*a_rate(t)
    time_period:      A value to be time_periodtituted for the time index, t.
    variables:        A list of all variables extracted from excel sheet.
    depth:            Depth of recursion, used internally
    returns:          A dict with a segment as key and computed excel cell index as value, e.g: {a(t - 1): (5, 4), a_rate(t): (4, 5)}
    """
    result = {}
    variable = expression.func        # get the function from sympy expression, e.g for expression = f(t), `f` is the function
    if variable.is_Function:
        # for simple expressions like f(t), variable=f and variable.is_Function = True,
        # for more complex expressions, variable would be another expression, hence would have to be broken down recursively.
        cell_row = variables[str(variable)]            # get the row index from variable name
        x = list(expression.args[0].free_symbols)[0]   # get the independent var, mostly `t` from the argument in expression
        cell_col = int(expression.args[0].subs(x, time_period))
        result[expression] = (cell_row, cell_col)
    else:
        if depth > 5:
            raise ValueError("Expression is too complicated: " + expression)

        depth += 1
        for segment in expression.args:
            result.update(simplify_expression(segment, time_period, variables, depth))

    return result

def get_excel_formula_as_string(right_side_expression, time_period, variables):
    """
    Using the right-hand side of a math expression (e.g. a(t)=a(t-1)*a_rate(t)), converted to sympy
    expression, and substituting the time index variable (t) in it, the function finds the Excel formula
    corresponding to the right-hand side expression.
    input
    -----
    right_side_expression:         sympy expression, e.g. a(t-1)*a_rate(t)
    time_period:        value of time index variable (t) for time_periodtitution
    output:
    formula_string:     a string of excel formula, e.g. '=A20*B21'
    """
    right_dict = simplify_expression(right_side_expression, time_period, variables)
    for right_key, right_coords in right_dict.items():
        excel_index = str(Range(get_sheet(), tuple(right_coords)).get_address(False, False))
        right_side_expression = right_side_expression.subs(right_key, excel_index)
    formula_str = '=' + str(right_side_expression)
    return formula_str

def _get_formula(parsed_formulas, row, col):
    """
    Returns the formula for a given row and column.
    """
    try:
        formula_dict = parsed_formulas[row]
    except KeyError:
        formula_dict = dict()
        if Range(get_sheet(), (row, col)).value is None:    # if cell is empty and formula for it not found
            print("Warning: Formula for empty cell not found, incomplete sheet, cell: " +
                  Range(get_sheet(), (row, col)).get_address(False, False))

    return formula_dict

def apply_formulas_on_sheet(workbook, variables, parsed_formulas, start_cell):
    """
    Takes each cell in the sheet inside the rectangle formed by Start_cell and End_cell
    checks 1) if the cell is in a row with a variable as first element
           2) if the cell is in a column with `is_forecast=1`
    If all above conditions are met, then apply a fitting formula as obtained from find_formulas()
    Apply's the solution on the workbook cells. Raises error if any problem arises.
    input
    -----
    workbook:   Workbook xlwings object
    variables: A dict of variables from excel sheet
    parsed_formulas: A dict of formulas with key as row_index and value as dict of left-side and right-side sympy expressions
    start_cell: Start cell dictionary
    """
    workbook.set_current()    # sets the workbook as the current working workbook
    forecast_row = Range(get_sheet(), (variables['is_forecast'], start_cell['col'] + 1)).horizontal.value
    col_indices = [start_cell['col'] + 1 + index for index, el in enumerate(forecast_row) if el == 1]    # checks if is_forecast value in this col is = 1 and notes down col index
    row_indices = list(variables.values())
    row_indices.remove(variables['is_forecast'])

    for col, row in itertools.product(col_indices, row_indices):

        formula_dict = _get_formula(parsed_formulas, row, col)

        if formula_dict:
            dependent_variable_with_time_index = formula_dict['dependent_var']     # get expression for dependent variable, e.g. a(t)
            # dependent_variable_locations - values like {b(t): (8, 5)}
            dependent_variable_locations = simplify_expression(dependent_variable_with_time_index, col, variables)
            dv_key, dv_coords = dependent_variable_locations.popitem()

            # 2015-05-12 03:09 PM
            # --- Need to make this check elsewhere
            if dependent_variable_locations:
                raise ValueError('cannot have more than one dependent variable on left side of equation')
            # --- end

            # find excel type formula string
            right_side_expression = formula_dict['formula']
            formula_str = get_excel_formula_as_string(right_side_expression, col, variables)
            Range(get_sheet(), dv_coords).formula = formula_str                # Apply formula on excel cell
