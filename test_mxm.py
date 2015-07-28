# -*- coding: utf-8 -*-
"""
Created on Wed Jul 29 00:25:34 2015

@author: Евгений
"""
import numpy as np
import os
from make_xl_model import get_resulting_workbook_array


def test_spec_xls_final_output():
   
    ref_ar = np.array([
      ['', '2013', '2014', '2015', '2016']
     ,['GDP', 66190, 71406, '=C2*D3/100*D4/100', '=D2*E3/100*E4/100']
     ,['GDP_IQ', 101.3, 100.6, 95.0, 102.5]
     ,['GDP_IP', 105.0, 107.2, 115.0, 113.0] ]
     , dtype=object)
                                  
    abs_filepath = os.path.abspath("spec.xls")    
    ar = get_resulting_workbook_array(abs_filepath) 
    assert np.array_equal(ar, ref_ar) is True


def test_spec_xls_final_output():
   
    ref_ar = np.array([
      ['', '2013', '2014', '2015', '2016']
     ,['GDP', 66190, 71406, '=C2*D3/100*D4/100', '=D2*E3/100*E4/100']
     ,['GDP_IQ', 101.3, 100.6, 95.0, 102.5]
     ,['GDP_IP', 105.0, 107.2, 115.0, 113.0] ]
     , dtype=object)
                                  
    abs_filepath = os.path.abspath("spec.xls")    
    ar = get_resulting_workbook_array(abs_filepath) 
    assert np.array_equal(ar, ref_ar) is True