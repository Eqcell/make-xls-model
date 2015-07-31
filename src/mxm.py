"""Make spreadsheet model in Excel file based on historic data, equations, and control parameters.
   Applies to forecasts of annual frequency time series.      
   
Usage:   
    mxm.py <xlfile> -M 
    mxm.py <xlfile> -M [--slim | -s]
    mxm.py <xlfile> -M [--fancy | -f]    
    mxm.py <xlfile> -U [--sheet=<sheet>]
"""

"""
Options:
   -M  Reads inputs from 'data', 'controls', 'equations' and 'names' sheets of <xlfile>
       and writes resulting spreadsheet to 'model' sheet in <xlfile>. Overwrites 'model'
       sheet in <xlfile>  without warning.
       
       Flags:
           --slim or -s produce minimum formatting on 'model' sheet (labels and years only).
           --fancy or -f produce extra formatting of 'model' sheet.    
           Default behaviour: fancy.
           
   -U  Updates Excel formulas on <sheet> only by reading this sheet. Works on output of 
       'mxm.py -M <xlfile> -f'. Default for <sheet> is 'model'.       
"""

from docopt import docopt
import os
from make_xl_model import make_xl_model, update_xl_model
from globals import MODEL_SHEET

def get_abs_filepath(arg):
   """Returns absolute path to <xlfile>"""
   return os.path.abspath(arg["<xlfile>"])
    
def get_model_sheet(arg):
   return MODEL_SHEET

if __name__ == "__main__":
   arg = docopt(__doc__)
    
   file = get_abs_filepath(arg)
   sheet = get_model_sheet(arg)    
   
   # default behaviour is slim formatting
   if arg["--slim"] or arg["-s"]:
       slim = True
   else:
       slim = False
   
   if arg["-U"]:
      # third column pivot is default output of -M --fancy keys
      update_xl_model(file, sheet, pivot_col = 2)
   elif arg["-M"]:
      make_xl_model(file, sheet, slim)
   else:
      raise Exception ("CLI input not specified.")
   