# ***************** new code *****************
from sympy import var
TIME_INDEX_VARIABLES = ['t', 'T', 'n', 'N']

def get_excel_ref(cell):  
    """
    TODO: test below fails with strange message, need to fix
    
    >>> get_excel_ref((0,0))
    'A1'
    
    >>> get_excel_ref((1,3))
    'D2'
    
    """
    # was xlwings's
    # str(Range(get_sheet(), tuple(right_coords)).get_address(False, False))
    # better (not todo): substitute with some xl package own formula engine or fork from there
    return get_xl_col_litteral(cell[1]) + str(cell[0]+1) 

def get_xl_col_litteral(zero_based_col_number):
    """
    Returns A...ZZZ type of string corresponding to *zero_based_col_number* col number
    >>> get_xl_col_litteral(0)
    'A'
    >>> get_xl_col_litteral(3-1)
    'C'
    """
    return "ABCDEFGHIJK"[zero_based_col_number]

def sympyfy_formula(string):
    return evaluate_variable(string)

def evaluate_variable(x):
    try:
        x = eval(x)     # converting the formula into sympy expressions
    except NameError:
        raise NameError('Undefined variables in formulas, check excel sheet')
    return x

def check_parse_equation_to_xl_formula():   
    """
    >>> check_parse_equation_to_xl_formula()
    =D2*E3*E4/10000
    """
    dict_formula = {'dependent_var': 'GDP(t)',
         'formula': 'GDP(t-1) * GDP_IP(t) / 100 * GDP_IQ(t) / 100'}
   
    # WARNING = actual dict_variables contains {'': 0,} - this is a bigger issue that it seems, to check later
    dict_variables = {'GDP': 1, 'GDP_IP': 2, 'GDP_IQ': 3}
    
    print (parse_equation_to_xl_formula(dict_formula, dict_variables, 4))

def get_cell_row(dict_variables, var_name):    
    return dict_variables[str(var_name)]


def parse_equation_to_xl_formula(formula_as_string, dict_variables, column):
    varirable_list = [x for x in dict_variables.keys()] + TIME_INDEX_VARIABLES

    # declares sympy variables
    var(' '.join(varirable_list))

   
    right_side_expression = sympyfy_formula(formula_as_string)
    time_period = column
    variables = dict_variables
    
    return get_excel_formula_as_string(right_side_expression, time_period, dict_variables)
    
    """
    must have somewhere:
        row, col - cell location
        x[t] = x[t-1] - equation
        x - variable name
        variable_locations dict rows for all other variables {var1:row1, var2:row2}
    require:
        left-hand side is always [t]. cannot accept [t-1], [t+1] or other on the left of '='
    algorithm:
        in equation we search for terms that denote variables, possibly lagged
        term = 'x(t-1)'
        *col* is substituted for *t* in 't-1' and evaluated, result is column number
        col = 5, index = 5-1 = 4
        variable_locations['x'] equals row with that variable
    returns:
        excel formula to be pasted to (row, col)

    """

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
        #excel_index = str(Range(get_sheet(), tuple(right_coords)).get_address(False, False))
        excel_index = get_excel_ref(tuple(right_coords))
        right_side_expression = right_side_expression.subs(right_key, excel_index)
    formula_str = '=' + str(right_side_expression)
    return formula_str

def simplify_expression(expression, time_period, variables, depth=0):
    """
    A recursive function which breaks a Sympy expression into segments,
    where each segment points to one cell on the excel sheet upon substitution
    of time index variable (t). Returns a dictionary of such segments and the computed
    cells.
    input
    -----
    expression:       Sympy expression, e.g: a(t - 1)*a_rate(t)
    time_period:      A value to be time_periodtituted for the time index, t.
    variables:        A list of all variables extracted from excel sheet.
    depth:            Depth of recursion, used internally

    returns:          A dict with a segment as key and computed excel cell index as value,
                      e.g: {a(t - 1): (5, 4), a_rate(t): (4, 5)}
    """
    result = {}

    # get the function from sympy expression, e.g for expression = f(t), `f` is the function
    variable = expression.func

    if variable.is_Function:
        # for simple expressions like f(t), variable=f and variable.is_Function = True,
        # for more complex expressions, variable would be another expression, hence would have to be broken down recursively.
        # get the row index from variable name
        # cell_row = variables[str(variable)]
        cell_row = get_cell_row(variables, variable)
        # get the independent var, mostly `t` from the argument in expression
        x = list(expression.args[0].free_symbols)[0]
        cell_col = int(expression.args[0].subs(x, time_period))
        result[expression] = (cell_row, cell_col)
    else:
        if depth > 5:
            raise ValueError("Expression is too complicated: " + expression)

        depth += 1
        for segment in expression.args:
            result.update(simplify_expression(segment, time_period, variables, depth))

    return result
    
    
        
           
           
if __name__ == "__main__":
    import doctest
    doctest.testmod()
    
    print("\n*** Sample formula:")
    check_parse_equation_to_xl_formula()


