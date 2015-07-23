'''Module to parse formulas to their respective excel representations'''

#duplicate
TIME_INDEX_VARIABLES = ['t', 'T', 'n', 'N']

"""EP:

Scope of work:
a) in this file -  obtain a text string representing *formula_as_string* based on cell locations defined by *variables_dict* 
   and *time_period*
   'GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100' ->  '=D2*E3/100*E4/100'
b) todos in https://github.com/epogrebnyak/make-xls-model/blob/master/equations_preparser.py (used in xl.fill 
   when splitting formulas)

c) proper eqcell_core.get_xl_col_litteral(zero_based_col_number):
   see eg  http://stackoverflow.com/questions/19415937/python-xlwt-convert-column-integer-into-excel-cell-references-eg-3-6-to-c6
   I think xlrd.colname() is safest, as xlrd appears a part of Anaconda installation.


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
    # TODO: formula_as_string will have a structure of this kind
    #       'GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100'
    #
    #       we want to substitute the variables with the appropriate excel
    #       values:
    #       =D2*E3 / 100*E4 / 100
    
    # TODO 1: for each variable build a regex that replaces
    #         make_regex('GDP_IQ', 't-1') -> <regex>
    #         make_regex('GDP', 't-1') -> <regex>
    # TODO 2: For each dict_variables, retrieve corresponding excel cell string.
    #         get_cell_name(dict_variables['GDP'], time_period) -> D1
    # TODO 3: For each regex, replace the variable with the cell string
    #         formula = <regex>.sub(formula) for each <regex>
    #         GDP[t-1] + GDP_IQ[t-1] -> D1 + D2
    pass

def check_parse_equation_as_formula():
    return (internal_parse_equation_as_formula() == 
            parse_equation_to_xl_formula('GDP[t-1] + GDP_IQ[t]/10 + 1 ',
                                         variables_dict={'GDP': 2, 'GDP_IQ': 3},
                                         time_period=1))
                                        
def internal_parse_equation_as_formula():
    return 'D3 + E4/10 + 1'

if __name__ == '__main__':
    print(check_parse_equation_as_formula())
