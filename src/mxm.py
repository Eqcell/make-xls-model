"""Make spreadsheet model in Excel file based on historic data, equations, and control parameters.      
   
Usage:   
    mxm.py make <xlfile> [--slim | -s]
    mxm.py make <xlfile> [ --fancy | -f]
    mxm.py -M <xlfile>   [--slim | -s]
    mxm.py -M <xlfile>   [ --fancy | -f]    
    mxm.py update <xlfile> [--sheet=<sheet>]
    mxm.py -U <xlfile> [--sheet=<sheet>]
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
   if arg["--fancy"] or arg["-f"]:
       slim = False
   else:
       slim = True
   
   if arg["update"] or arg["-U"]:
      # third column pivot is default output of -M --fancy keys
      update_xl_model(file, sheet, pivot_col = 2)
   elif arg["make"] or arg["-M"]:
      make_xl_model(file, sheet, slim)
   else:
      raise Exception ("CLI input not specified.")
   