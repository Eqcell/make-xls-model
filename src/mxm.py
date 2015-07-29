"""Make spreadsheet model in Excel file based on historic data, equations, and control parameters.      
   
Usage:   
    mxm.py <xlfile> [--make | --update] [--slim | --fancy]
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
   
   if arg["--fancy"]:
       slim = False
   else:
       slim = True
   
   if arg["--update"]:
      update_xl_model(file, sheet, slim)
   elif arg["--make"]:
      make_xl_model(file, sheet, slim)
   