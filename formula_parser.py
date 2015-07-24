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

def parse_equation_to_xl_formula(formula_as_string, variables_dict, time_period):
    '''Equivalent method of eqcell_core, but with text-based parser
    
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
    issue a warning
    
    >>> parse_equation_to_xl_formula('GDP[t] + GDP_IQ[t-1] * 100',
    ...                              {'GDP': 1, 'GDP_IQ': 2, 'GDP_IP': 3}, 1)
    WARNING: Variable 'GDP_IP' included in variables_dict but not present in formula
    '=B2+A3*100'

    '''
    # Strip whitespace
    formula_as_string = re.sub(r'\s+', '', formula_as_string)

    # Extract a dictionary containing the pairs variable, period from
    # formula_as_string
    # variables_period = {'GDP' : 't-1', GDP_IP: 't'}
    variables_periods = extract_variables_periods(formula_as_string)
    
    # For each variable, substitute each VAR_NAME[PERIOD] with the 
    # corresponding excel cell name
    
    all_vars = set(variables_dict.keys()) # Check if all variables are consumed
    for var, period in variables_periods.items():
        # Calculate row, column of excel cell
        try:
            var_offset = variables_dict[var] # Variable offset in file
        except KeyError:
            raise ValueError('Variable %s included in formula should be included in variables_dict' % repr(var))
        
        period_normalize = re.sub('[' + ''.join(TIME_INDEX_VARIABLES) + ']', 't', period)
        t = time_period                            # Offset for reference period
        try:
            # We transfrom TIME_INDEX_VARIABLES to t for proper evaluation
            period_offset = eval(period_normalize)     # evaluate time expression t-1
        except:
            raise ValueError('Time expression %s[%s] invalid' % (var, period))
        
        # Get excel cell as string
        cell_string = get_excel_ref((var_offset, period_offset))
        
        # Build regular expression that will match VAR_NAME[PERIOD]
        regex = make_regex(var, period)
        # Perform actual substutution with cell string
        formula_as_string = regex.sub(cell_string, formula_as_string)
        all_vars.remove(var)
    
    if len(all_vars) != 0:
        for var in all_vars:
            if var != '':
                print('WARNING: Variable %s included in variables_dict but not present in formula' % repr(var))
    
    return '=' + formula_as_string

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


def expand_shortand(formula_as_string, variables):
    '''
    >>> expand_shortand('GDP_IQ+GDP_IP[t]+GDP_IQ[t-1]', {'GDP_IP': 1, 'GDP_IQ': 1})
    'GDP_IQ[t]+GDP_IP[t]+GDP_IQ[t-1]'
    '''
    for var in variables:
        formula_as_string = re.sub(var + r'(?!\[)', var + '[t]', formula_as_string) 
    return formula_as_string

def make_regex(var_name, period):
    '''Make regex to match var_name and period
    >>> make_regex('GDP', 't')
    re.compile('GDP\\\\[t\\\\]')
    '''
    return re.compile(r'%s\[%s\]' % (var_name, period))


def extract_variables_periods(formula_as_string):
    '''Extract variables from formula with their respective periods
    
    >>> extract_variables_periods('GDP[t-1]+GDP_IQ[t]') == {'GDP_IQ': 't', 'GDP': 't-1'}
    True
    '''
    # Extract the variable expressions  in a list ['GDP[t-1]', 'GDP_IQ[t]']
    variable_time_dep = re.findall(r'\w+\[[' + ''.join(TIME_INDEX_VARIABLES) + '+\-\d]+\]', formula_as_string)
    
    # Extract groups [(GDP, t-1), (GDP_IQ, t)]
    variable_time_dep_grouped = [re.match(r'(\w+)\[(.+)\]', v).groups() 
                                              for v in variable_time_dep]
    return dict(variable_time_dep_grouped)

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
