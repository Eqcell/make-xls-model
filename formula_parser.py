'''Module to parse formulas to their respective excel representations'''
import re
from eqcell_core import get_excel_ref


def parse_equation_to_xl_formula(formula_as_string, variables_dict, time_period):
    '''Equivalent method of eqcell_core, but with text-based parser'''
    # Extract a dictionary containing the pairs variable, period from
    # formula_as_string
    # variables_period ~ {'GDP' : 't-1', GDP_IP: 't'}
    variables_periods = extract_variables_periods(formula_as_string)
    
    # For ecah variable, substitute each VAR_NAME(PERIOD) with the 
    # corresponding excel cell name
    for var, period in variables_periods.items():
        
        # Build regular expression that will match VAR_NAME(PERIOD)
        regex = make_regex(var, period)
        
        # Retrieve excel cell
        
        var_offset = variables_dict[var] # Variable offset in file
        t = time_period                  # Offset for reference period
        period_offset = eval(period)     # evaluate time expression t-1
        
        # Get cell string
        cell_string = get_excel_ref((var_offset, period_offset))
        
        # Perform actual substutution with cell string
        formula_as_string = regex.sub(cell_string, formula_as_string)

    return '=' + formula_as_string

def get_excel_ref_for_var_period(variable, period, var_offset, time_offset):
    t = time_offset
    return get_excel_ref((var_offset, eval(period)))


def make_regex(var_name, period):
    return re.compile(r'%s\s*\(%s\)' % (var_name, period))


def extract_variables_periods(formula_as_string):
    # Extract GDP(t-1), GDP_IQ(t)
    variable_time_dep = re.findall(r'\w+\([t+\-\d]+\)', formula_as_string)
    # Extract groups [(GDP, t-1), (GDP_IQ, t)]
    variable_time_dep_grouped = [re.match(r'(\w+)\((.+)\)', v).groups() 
                                              for v in variable_time_dep]
    return dict(variable_time_dep_grouped)

def check_parse_equation_as_formula():
    """
    >>> check_parse_equation_as_formula()
    '=C2 * D4 / 100 * D3 / 100'
    
    """
    formula_as_string = 'GDP(t-1) * GDP_IP(t) / 100 * GDP_IQ(t) / 100'
    variables_dict = {'': 0, 'GDP_IQ': 2, 'GDP': 1, 'GDP_IP': 3}
    time_period = 3
    return parse_equation_to_xl_formula(formula_as_string, variables_dict, time_period)

if __name__ == '__main__':
    import doctest
    doctest.testmod()