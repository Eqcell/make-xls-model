"""
   Parsing of formulas as strings into dictionary. Used by either formula parser.  
"""

import re
    
def strip_timeindex(str_):
    """Returns variable name without time index.
    
       Accepted *str_*: 'GDP', 'GDP[t]', 'GDP(t)', '    GDP [ t ]', ' GDP   ( t) '       
    
    TODO: must work both with [t] and t()
          must accept variable names without brackets
          must accept whitespace anywhere
          (see failed tests)
   
    Passed test:
    >>> strip_timeindex("GDP(t)")
    'GDP'
        
    Failing tests (4 tests):    
    >>> strip_timeindex("GDP[t]")
    'GDP'

    >>> strip_timeindex("GDP")
    'GDP'
    
    >>> strip_timeindex('    GDP [ t ]')
    'GDP'
     
    >>> strip_timeindex(' GDP   ( t) ')
    'GDP'
    """
    if "[" in str_ or "(" in str_:
        pattern = r"\s*(\S*)\s*[(\[].*[)\]]"
        m = re.search(pattern, str_)
        if m:
            return m.groups()[0]
        else:
            raise ValueError('Error extracting variable names from: ' + str_)
    else:
        return str_.strip()
             

    
def test_parse_to_formula_dict():    
    """
    >>> test_parse_to_formula_dict()
    True
    True
    True
    True
    """
    inputs = [
      ['GDP(t) = GDP(t-1) * GDP_IP(t) / 100 * GDP_IQ(t) / 100']
    , ['x(t) = x(t-1) + 1']
    , ['x(t) = x(t-1) + 1', 'y(t) = x(t)']
    , ['credit = credit[t-1] * credit_rog'] 
    ]    
    expected_outputs = [
      {'GDP': ['GDP(t)', 'GDP(t-1) * GDP_IP(t) / 100 * GDP_IQ(t) / 100']}   
    , {'x':   ['x(t)', 'x(t-1) + 1']                                    }
    , {'x':   ['x(t)', 'x(t-1) + 1'], 'y': ['y(t)', 'x(t)']             }
    , {'credit' : ['credit', 'credit[t-1] * credit_rog']                } 
    ]
    for input_eq, expected_output in zip(inputs,expected_outputs):
       print(expected_output == parse_to_formula_dict(input_eq))

def parse_to_formula_dict(equations_list_of_strings):
    """Returns a dict with left and right hand side of equations, referenced by variable name in keys."""
    parsed_eq_dict = {}
    for eq in equations_list_of_strings:
        dependent_var, formula = eq.split('=')
        key = strip_timeindex(dependent_var)
        # parsed_eq_dict[key] = {'dependent_var': dependent_var.strip(), 'formula': formula.strip()}
        parsed_eq_dict[key] = [dependent_var.strip(), formula.strip()]
    return parsed_eq_dict

def get_formula(var_name, eq_dict):
    """Returns a formula for *var_name* based on contents of *eq_dict*.
    
    Test:
    >>> get_formula('x', {'x':[None, 'x+1']})
    'x+1'
    """
    return eq_dict[var_name][1]
    
if __name__ == "__main__":   
    import doctest
    doctest.testmod()
