'''Module to parse formulas to their respective excel representations'''
import re
import xlrd
from config import TIME_INDEX_VARIABLES

"""EP:

Outstanding: 
   a) and e)
   
Scope of work:
a) in this file -  obtain a text string representing *formula_string* based on cell locations defined by *variables_dict* 
   and *time_period*
   'GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100' ->  '=D2*E3/100*E4/100'
   EP: implemented, but error in https://github.com/epogrebnyak/make-xls-model/issues/9
   
b) todos in https://github.com/epogrebnyak/make-xls-model/blob/master/equations_preparser.py (used in xl.fill 
   when splitting formulas)
   EP: now done.

c) proper eqcell_core.get_xl_col_litteral(zero_based_col_number):
   see eg  http://stackoverflow.com/questions/19415937/python-xlwt-convert-column-integer-into-excel-cell-references-eg-3-6-to-c6
   I think xlrd.colname() is safest, as xlrd appears a part of Anaconda installation.
   EP: done
   
d) eqcell_core.py - cosmetic change (will be depreciated)
   EP: under way.    
   
e) placing TIME_INDEX_VARIABLES somewhere to avoid corss-reference between the files. Maybe even a new config.py
   EP: done

Expected benefit:
   formula parser not dependent on sympy package
   
Assumptions:
  - time index in square brackets [], [<ws><litteral in TIME_INDEX_VARIABLES><ws >+-<ws><integer><ws>]
  <ws> is whitespace, any or none
  - any variable from variables_dict would appear in formula with a time index in brackets 
  - variable is detected as participating in variables_dict.keys()

Suggested implementation:
  - strip all whitespace from inside of formula_string (it would simplyfy the rest of parsing and spaces 
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

def parse_equation_to_xl_formula(formula_string, variables_dict, time_period):
    '''Equivalent method of eqcell_core, but with text-based parser

    >>> parse_equation_to_xl_formula('liq_to_credit*credit', {'credit':10, 'liq_to_credit': 9}, 1)
    '=B10*B11'
    
    >>> parse_equation_to_xl_formula('liq_to_credit*credit', {'liq_to_credit': 9, 'credit':10}, 1)
    '=B10*B11'
        
    >>> parse_equation_to_xl_formula('GDP[t]', {'GDP': 99}, 1)
    '=B100'
    
    >>> parse_equation_to_xl_formula('GDP[t] * 0.5 + GDP[t-1] * 0.5',
    ...                              {'GDP': 99}, 1)
    '=B100*0.5+A100*0.5'
    
    >>> parse_equation_to_xl_formula('GDP * 0.5 + GDP[t-1] * 0.5',
    ...                              {'GDP': 99}, 1)
    '=B100*0.5+A100*0.5'

    >>> parse_equation_to_xl_formula('liq[t] + credit[t] * 0.5 + liq_to_credit[t] * 0.5',
    ...                              {'credit': 2, 'liq_to_credit': 3, 'liq': 4}, 1)
    '=B5+B3*0.5+B4*0.5'
    
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
    KeyError: Cannot parse formula, formula contains unknown variable: GDP_IQ
    
    If some variable is included in variables_dict but do not appear in formula_string
    do nothing.
    
    >>> parse_equation_to_xl_formula('GDP[t] + GDP_IQ[t-1] * 100',
    ...                              {'GDP': 1, 'GDP_IQ': 2, 'GDP_IP': 3}, 1)
    '=B2+A3*100'

    '''
    # Strip whitespace
    formula_string = strip_all_whitespace(formula_string)
    
    # Expands shorthand
    formula_string = expand_shorthand(formula_string, variables_dict.keys())
    
    # parse and substitute time indices, eg. GDP[t-1] -> GDP[3] if t = 4
    formula_string = substitute_time_indices(formula_string, time_period)
    
    # each setment in var_time_segments  is like 'GDP[0]', 'GDP_IQ[10]', etc
    var_time_segments = re.findall(r'(\w+\[\d+\])', formula_string)
    for segment in var_time_segments: 
        formula_string = replace_segment_in_formula(formula_string, segment, variables_dict)
        
    return '=' + formula_string

def strip_all_whitespace(string):
    return re.sub(r'\s+', '', string)
    
def get_A1_reference(segment, variables_dict):
    var, period = extract_var_time(segment)
    if var in variables_dict.keys():
        cell_row = get_cell_row(var, variables_dict)
        cell_col = period  
        return get_excel_ref(cell_row, period) 
    else:
        raise KeyError("Cannot parse formula, formula contains unknown variable: " + var)        
    
def replace_segment_in_formula(formula_string, segment, variables_dict):
    A1_ref = get_A1_reference(segment, variables_dict)
    # Match beginning of word
    return re.sub(r'\b' + re.escape(segment), A1_ref, formula_string)
    
def get_cell_row(var, variables_dict):
    try:
        return variables_dict[var] # Variable offset in file
    except KeyError:
        raise ValueError('Variable %s is in formula, but not found in variables_dict' % repr(var))

def get_excel_ref(row, col):
    '''
    >>> get_excel_ref(0, 0)
    'A1'
    >>> get_excel_ref(3, 2)
    'C4'
    '''
    return xlrd.colname(col) + str(row + 1)

def substitute_time_indices(formula_string, period):
    '''
    >>> substitute_time_indices('GDP[t]+GDP[t-1]+0.5*GDP_IP[t]', 1)
    'GDP[1]+GDP[0]+0.5*GDP_IP[1]'
    >>> substitute_time_indices('GDP[t]+GDP[n-1]+0.5*GDP_IP[n]', 1)
    'GDP[1]+GDP[0]+0.5*GDP_IP[1]'
    '''
    
    # time index in square brackets [], [<ws><litteral in TIME_INDEX_VARIABLES><ws >+-<ws><integer><ws>]
    TI = ''.join(TIME_INDEX_VARIABLES)
    # note here [] are part of regex notation 
    TI_REGEX = r'[' + TI + r'+\-\d]'
    
    for time_index in re.findall(r'\[(' + TI_REGEX + '+)\]', formula_string):
        # We transfrom TIME_INDEX_VARIABLES to t for proper evaluation
        period_normalize = re.sub('[' + TI + ']', 't', time_index)
        try:
            t = period
            period_offset = eval(period_normalize)     # evaluate time expression t-1
        except:
            raise ValueError('Time expression %s[%s] invalid' % (var, period))
        
        formula_string = formula_string.replace('[' + time_index + ']', 
                                                      '[' + str(period_offset) + ']')
    return formula_string

def expand_shorthand(formula_string, variables):
    """
    >>> expand_shorthand('GDP_IQ+GDP_IP+GDP_IQ[t-1]', ['GDP', 'GDP_IP', 'GDP_IQ'])
    'GDP_IQ[t]+GDP_IP[t]+GDP_IQ[t-1]'
    
    >>> expand_shorthand('GDP * 0 + GDP [t-1] * GDP_IQ / 100 * GDP_IP[t] / 100', 
    ...                        ['GDP', 'GDP_IP', 'GDP_IQ'])
    'GDP[t] * 0 + GDP [t-1] * GDP_IQ[t] / 100 * GDP_IP[t] / 100'
    
    >>> expand_shorthand('GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100',
    ...                  ['GDP', 'GDP_IP', 'GDP_IQ'])
    'GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100'
    
    >>> expand_shorthand('liq_to_credit*credit', ['liq_to_credit', 'credit'])
    'liq_to_credit[t]*credit[t]'
    
    >>> expand_shorthand('liq_to_credit*credit', ['credit', 'liq_to_credit'])
    'liq_to_credit[t]*credit[t]'
    
    """
    for var in variables:
        if var != '':
            formula_string = re.sub(var + r'(?!\s*[\dA-Za-z_^\[])',
                                       var + '[t]', formula_string) 
    return formula_string

def make_regex(var_name, period):
    '''Make regex to match var_name and period    
    >>> make_regex('GDP', 't')
    re.compile('GDP\\\\[t\\\\]')
    '''
    return re.compile(r'%s\[%s\]' % (var_name, period))

def extract_var_time(formula_string):
    '''Extract variable and time period from formula segment
    
    >>> extract_var_time('GDP[1]')
    ('GDP', 1)
    '''
    # Extract group (GDP, 0)
    a, b = re.search(r'(\w+)\[(\d+)\]', formula_string).groups()
    return a, int(b)

def check_parse_equation_as_formula():
    """
    >>> check_parse_equation_as_formula()
    '=C2*D4/100*D3/100'
    
    """
    formula_string = 'GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100'
    variables_dict = {'': 0, 'GDP_IQ': 2, 'GDP': 1, 'GDP_IP': 3}
    time_period = 3
    return parse_equation_to_xl_formula(formula_string, variables_dict, time_period)

if __name__ == '__main__':
    import doctest
    doctest.testmod()
    pass
    
