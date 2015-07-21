'''Module to parse formulas to their respective excel representations'''

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