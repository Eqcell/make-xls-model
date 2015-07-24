'''Module to parse formulas to their respective excel representations'''
import re
import xlrd

#duplicate
TIME_INDEX_VARIABLES = ['t', 'T', 'n', 'N']

"""EP:

Done:
   I moved splitting of text string to equations_preparser.py

Scope of work:
a) in this file -  obtain a text string representing *formula_as_string* based on cell locations defined by *variables_dict* 
   and *time_period*
   'GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100' ->  '=D2*E3/100*E4/100'
   
b) todos in https://github.com/epogrebnyak/make-xls-model/blob/master/equations_preparser.py (used in xl.fill 
   when splitting formulas)

c) proper eqcell_core.get_xl_col_litteral(zero_based_col_number):
   see eg  http://stackoverflow.com/questions/19415937/python-xlwt-convert-column-integer-into-excel-cell-references-eg-3-6-to-c6
   I think xlrd.colname() is safest, as xlrd appears a part of Anaconda installation.

d) eqcell_core.py - cosmetic change (will be depreciated)

e) placing TIME_INDEX_VARIABLES somewhere to avoid corss-reference between the files. Maybe even a new config.py

Expected benefit:
   formula parser not dependent on sympy package
   
Assumptions:
  - time index in square brackets [], [<ws><litteral in TIME_INDEX_VARIABLES><ws >+-<ws><integer><ws>]
  <ws> is whitespace, any or none
  - any variable from variables_dict would appear in formula with a time index in brackets 
  - variable is detected as participating in variables_dict.keys()

Suggested implementation:
  - strip all whitespace from inside of formula_as_string (it would simplyfy the rest of parsing and spaces 
    get eaten up in Excel anyways)
  - parse and substitute  time indices, egg. GDP[t-1] -> GDP[3] if t = 4
  - for each in variables_dict.keys() change (<varname>)\[(<int>)\] to xl A1 reference 
  
Validation: 
  - if any brackets left - raise exception and print (formula contains variables not in variables_dict.keys() )
  - something else?
  
Testing:
    variables_dict={'GDP': 2, 'GDP_IQ': 3}, time_period=1
    # i think when period is time_period=1, column to [t] is B, as we see time index zero based (it is actually just a column number)
   ('GDP[t-1] + GDP_IQ[t]/10 + 1 ','A3+B4/10+1') # Note stripped whitespace
   ('GDP[t-1] + GDP_IQ/10 + 1 ','A3+B4/10+1')    # Shorthand  
   
                                  
                                        
def internal_parse_equation_as_formula():
    return 'D3 + E4/10 + 1'
  - maybe compare outputs with sypmy parser, for a check.
  
Additional behaviour: 
  - I want to include 'GDP' without time index as a shorthand notations for 'GDP(t)'
    this should be a valid formula: GDP = GDP[t-1] * ROG1 -> GDP[t] = GDP[t-1] * ROG1[t] 
    (ROG, rate of growth)
    
    

"""

def strip_all_whitespace(string):
    return re.sub(r'\s+', '', string)
    

def parse_equation_to_xl_formula(formula_as_string, variables_dict, time_period):
    '''Equivalent method of eqcell_core, but with text-based parser
    
    >>> parse_equation_to_xl_formula('GDP[t] * 0.5 + GDP[t-1] * 0.5',
    ...                              {'GDP': 99}, 1)
    '=B100*0.5+A100*0.5'
    
    >>> parse_equation_to_xl_formula('GDP * 0.5 + GDP[t-1] * 0.5',
    ...                              {'GDP': 99}, 1)
    '=B100*0.5+A100*0.5'
    
    >>> parse_equation_to_xl_formula('GDP[t] + GDP_IQ[t-1] * 100',
    ...                              {'GDP': 1, 'GDP_IQ': 2}, 1)
    '=B2+A3*100'
    
    >>> parse_equation_to_xl_formula('GDP[n] + GDP_IQ[n-1] * 100',
    ...                              {'GDP': 1, 'GDP_IQ': 2}, 1)
    '=B2+A3*100'
    
    If some variable is missing from 'variable_dict' raise an exception:
    
    >>> parse_equation_to_xl_formula('GDP[t] + GDP_IQ[t-1] * 100', # doctest: +IGNORE_EXCEPTION_DETAIL 
    ...                              {'GDP': 1}, 1)
    Traceback (most recent call last):  
    ValueError: Variable 'GDP_IQ' included in formula should be included in variables_dict
    
    If some variable is included in variables_dict but do not appear in formula_string
    do nothing.
    
    >>> parse_equation_to_xl_formula('GDP[t] + GDP_IQ[t-1] * 100',
    ...                              {'GDP': 1, 'GDP_IQ': 2, 'GDP_IP': 3}, 1)
    '=B2+A3*100'

    '''
    # Strip whitespace
    formula_as_string = strip_all_whitespace(formula_as_string)
    
    # Expands shorthand
    formula_as_string = expand_shorthand(formula_as_string, variables_dict.keys())
    
    # parse and substitute  time indices, egg. GDP[t-1] -> GDP[3] if t = 4
    formula_as_string = substitute_time_indices(formula_as_string, time_period)
    
    # Extract a list containing the pairs variable, period from
    # formula_as_string
    # variables_period = [('GDP', 0), (GDP_IP, 1)]
    variables_periods = extract_variables_periods(formula_as_string)
    
    # For each variable, substitute each VAR_NAME[PERIOD] with the 
    # corresponding excel cell name
    for var, period in variables_periods:
        # Calculate row, column of excel cell
        var_offset = get_cell_row(var, variables_dict)
        # Get excel cell as string
        cell_string = get_excel_ref((var_offset, period))
        
        # change (<varname>)\[(<int>)\] to xl A1 reference 
        formula_as_string = formula_as_string.replace('%s[%d]' % (var, period), cell_string)
    
    if '[' in formula_as_string or ']' in formula_as_string:
        raise ValueError('Formula contains variables not in %s' % variables_dict.keys())

    return '=' + formula_as_string

def get_cell_row(var, variables_dict):
    try:
        return variables_dict[var] # Variable offset in file
    except KeyError:
        raise ValueError('Variable %s included in formula should be included in variables_dict' % repr(var))

def get_excel_ref(cell):
    '''
    >>> get_excel_ref((0, 0))
    'A1'
    >>> get_excel_ref((3, 2))
    'C4'
    '''
    row, col = cell
    return xlrd.colname(col) + str(row + 1)

def get_excel_ref_for_var_period(variable, period, var_offset, time_offset):
    t = time_offset
    return get_excel_ref((var_offset, eval(period)))

def substitute_time_indices(formula_as_string, period):
    '''
    >>> substitute_time_indices('GDP[t]+GDP[t-1]+0.5*GDP_IP[t]', 1)
    'GDP[1]+GDP[0]+0.5*GDP_IP[1]'
    >>> substitute_time_indices('GDP[t]+GDP[n-1]+0.5*GDP_IP[n]', 1)
    'GDP[1]+GDP[0]+0.5*GDP_IP[1]'
    '''
    
    # time index in square brackets [], [<ws><litteral in TIME_INDEX_VARIABLES><ws >+-<ws><integer><ws>]
    TI = ''.join(TIME_INDEX_VARIABLES)
    TI_REGEX = r'[' + TI + r'+\-\d]'
    
    for time_index in re.findall(r'\[(' + TI_REGEX + '+)\]', formula_as_string):
        # We transfrom TIME_INDEX_VARIABLES to t for proper evaluation
        period_normalize = re.sub('[' + TI + ']', 't', time_index)
        try:
            t = period
            period_offset = eval(period_normalize)     # evaluate time expression t-1
        except:
            raise ValueError('Time expression %s[%s] invalid' % (var, period))
        
        formula_as_string = formula_as_string.replace('[' + time_index + ']', 
                                                      '[' + str(period_offset) + ']')
    return formula_as_string

def expand_shorthand(formula_as_string, variables):
    """
    >>> expand_shorthand('GDP_IQ+GDP_IP+GDP_IQ[t-1]', {'GDP_IP': 1, 'GDP_IQ': 2})
    'GDP_IQ[t]+GDP_IP[t]+GDP_IQ[t-1]'
    
    >>> expand_shorthand('GDP * 0 + GDP [t-1] * GDP_IQ / 100 * GDP_IP[t] / 100', {'GDP_IP': 1, 'GDP_IQ': 2, 'GDP':3})
    'GDP[t] * 0 + GDP [t-1] * GDP_IQ[t] / 100 * GDP_IP[t] / 100'
    
    >>> expand_shorthand('GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100', {'': 0, 'GDP_IQ': 2, 'GDP': 1, 'GDP_IP': 3})
    'GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100'
    """
    for var in variables:
        if var != '':
            formula_as_string = re.sub(var + r'(?!\s*[\dA-Za-z_^\[])',
                                       var + '[t]', formula_as_string) 
    return formula_as_string

def make_regex(var_name, period):
    '''Make regex to match var_name and period    
    >>> make_regex('GDP', 't')
    re.compile('GDP\\\\[t\\\\]')
    '''
    return re.compile(r'%s\[%s\]' % (var_name, period))


def extract_variables_periods(formula_as_string):
    '''Extract variables from formula with their respective periods
    
    >>> extract_variables_periods('GDP[1]+GDP_IQ[0]')
    [('GDP', 1), ('GDP_IQ', 0)]
    '''
    # Extract groups [(GDP, 0), (GDP_IQ, 1)]
    return [(a, int(b)) for a, b in re.findall(r'(\w+)\[(\d+)\]', formula_as_string)]


def check_parse_equation_as_formula():
    """
    >>> check_parse_equation_as_formula()
    '=C2*D4/100*D3/100'
    
    """
    formula_as_string = 'GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100'
    variables_dict = {'': 0, 'GDP_IQ': 2, 'GDP': 1, 'GDP_IP': 3}
    time_period = 3
    return parse_equation_to_xl_formula(formula_as_string, variables_dict, time_period)

if __name__ == '__main__':
    import doctest
    doctest.testmod()
    check_parse_equation_as_formula()
