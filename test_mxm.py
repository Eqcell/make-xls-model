# -*- coding: utf-8 -*-
"""Test suite for make_xls_model"""

import numpy as np
import pandas as pd
import os

xl_file = os.path.abspath("spec.xls") 

###########################################################################
## Data input
###########################################################################

def test_data_import():    
    from make_xl_model import get_data_df
    df =  get_data_df(xl_file)           
    df0 = pd.DataFrame(
             # Note we use float value for GDP, otherwise df.equals(df0) will fail
             { "GDP": [66190.0, 71406.0]
          , "GDP_IQ": [101.3407, 100.6404]       
          , "GDP_IP": [105.0467, 107.1941]}
          ,   index = [2013, 2014])[["GDP", "GDP_IQ", "GDP_IP"]] 
    assert df.equals(df0)

def test_controls_import():    
    from make_xl_model import get_controls_df
    df =  get_controls_df(xl_file)           
    df0 = pd.DataFrame(             
             {"GDP_IQ": [95.0, 102.5]       
            , "GDP_IP": [115.0, 113.0]}
            , index = [2015, 2016])[["GDP_IQ", "GDP_IP"]] 
    assert df.equals(df0)

def test_equation_import():     
    from make_xl_model import get_equations_dict
    ed = get_equations_dict(xl_file)                  
    assert ed == {'GDP': 'GDP[t-1] * GDP_IQ / 100 * GDP_IP / 100'}

###########################################################################
## Equations tests
########################################################################### 

def test_misspecified_equations():
    pass

def test_eq_with_comment():
    pass

def test_eq_without_eq_signt():
    pass

def test_imported_test():
    from formula_parser import test_make_eq_dict
    test_make_eq_dict()

def test_parse_equation_to_xl_formula():
    from formula_parser import parse_equation_to_xl_formula as pfunc
    
    time_period = 1
    
    dict_ = {'credit':10, 'liq_to_credit': 9}    
    assert pfunc('liq_to_credit*credit', dict_, 1) == '=B10*B11'
    
    dict_ = {'GDP': 99}   
    assert pfunc('GDP[t]', dict_, 1) == '=B100'
    assert pfunc('GDP'   , dict_, 1) == '=B100'    
    assert pfunc('GDP[t] * 0.5 + GDP[t-1] * 0.5', dict_, 1) == '=B100*0.5+A100*0.5'
    
    # todo:    
    
    '''
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
    '''
    
    # If some variable is included in variables_dict but do not appear in formula_string
    # do nothing.
    assert pfunc('GDP[t] + GDP_IQ[t-1] * 100',
           {'GDP': 1, 'GDP_IQ': 2, 'GDP_IP': 3}, 1) == '=B2+A3*100'

###########################################################################
## Quality of input
########################################################################### 

# Targets testing of:
# def validate_input_from_sheets(abs_filepath):
# def validate_continious_year(data_df, controls_df):
# def validate_coverage_by_equations(var_group, equations_dict):   

def test_data_not_covered_by_equations():
    pass

def years_continious():
    pass
    
###########################################################################
## Final result
###########################################################################    
    
def test_resulting_array_spec_xls():
    from make_xl_model import get_resulting_workbook_array    
    ar0 = np.array([
      ['', '2013', '2014', '2015', '2016']
     ,['GDP', 66190, 71406, '=C2*D3/100*D4/100', '=D2*E3/100*E4/100']
     ,['GDP_IQ', 101.3407,  100.6404, 95.0,  102.5]
     ,['GDP_IP', 105.0467,  107.1941, 115.0, 113.0] ]
     , dtype=object)              
    ar = get_resulting_workbook_array(xl_file)
    assert np.array_equal(ar, ar0) 
    
# also need test for this:
"""\nArray to write to Excel sheet:
[['' '2014' '2015' '2016' '2017' '2018']
 ['credit' 115.0 '=B2*C12' '=C2*D12' '=D2*E12' '=E2*F12']
 ['liq' 20.0 '=C2*C17' '=D2*D17' '=E2*E17' '=F2*F17']
 ['capital' 30.0 '=B4' '=C4' '=D4' '=E4']
 ['deposit' 90.0 '=B5*C13' '=C5*D13' '=D5*E13' '=E5*F13']
 ['profit' 10.0 '=C9*C15-C11*C16+C3*C14' '=D9*D15-D11*D16+D3*D14'
  '=E9*E15-E11*E16+E3*E14' '=F9*F15-F11*F16+F3*F14']
 ['acc_profit' 5.0 '=B7+B6' '=C7+C6' '=D7+D6' '=E7+E6']
 ['ta' '=B2+B3' '=C2+C3' '=D2+D3' '=E2+E3' '=F2+F3']
 ['avg_credit' '=0.5*B2+0.5*A2' '=0.5*C2+0.5*B2' '=0.5*D2+0.5*C2'
  '=0.5*E2+0.5*D2' '=0.5*F2+0.5*E2']
 ['fgap' '=B8-B4-B6-B5-B7' '=C8-C4-C6-C5-C7' '=D8-D4-D6-D5-D7'
  '=E8-E4-E6-E5-E7' '=F8-F4-F6-F5-F7']
 ['avg_deposit' '=0.5*B5+0.5*A5' '=0.5*C5+0.5*B5' '=0.5*D5+0.5*C5'
  '=0.5*E5+0.5*D5' '=0.5*F5+0.5*E5']
 ['credit_rog' nan 1.15 1.15 1.15 1.15]
 ['deposit_rog' nan 1.1 1.1 1.1 1.1]
 ['liq_ir' nan 0.02 0.02 0.02 0.02]
 ['credit_ir' nan 0.12 0.12 0.12 0.12]
 ['deposit_ir' nan 0.06 0.06 0.06 0.06]
 ['liq_to_credit' nan 0.2 0.2 0.2 0.2]]
"""
