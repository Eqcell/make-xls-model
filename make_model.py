# coding: cp1251
"""
   Generate Excel file with ordered rows containing Excel formulas 
   that allow to calculate forecast values based on historic data, 
   equations and forecast parameters. Order of rows in Excel file 
   controlled by template definition. Start year specified as input.

   Input:  
        data
        equations
        names
        controls (forecast parameters)
        format
        start_year        
   Output: 
        macro.xls
"""

import numpy as np

        

# todo: interface to eqcell.py
        
#***********************************************        

YEAR_ROW = 2 
LABEL_COLUMN = 3
CELL_TYPES = {'value':0, 'formula':1}

def


def create_empty_array(format):   
   max_rows = len(format['rows']) + 1
   max_col = 2 +(get_max_control_year() - format['start_year'] + 1)    
   return np.empty((max_rows, max_col), dtype = object)

range_array = create_empty_array(format)

# row = YEAR_ROW - 1,
# col = LABEL_COLUMN - 1
 
 
   
for j, label in enumerate(format['rows']):
    range_array[1 + j, 1] = label     #  (label, CELL_TYPES['value'])
    range_array[1 + j, 0] = 'Label text here' # label #  ('Label text here', CELL_TYPES['value'])  
print (range_array) 

#***********************************************    





