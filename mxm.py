"""Make Excel file with spreadsheet model based on historic data, equations, 
   control variables and spreadsheet parameters.      
   
Usage:   
   mxm.py --test
   mxm.py <XLFILE> [--markup <YAML_FILE>]
"""

from docopt import docopt
import os
from make_xl_model import make_xl_model

def get_abs_filepath(arg):
    """Returns absolute path to <XLFILE>"""
    return os.path.abspath(arg["<XLFILE>"])
    
def get_model_sheet(arg):
    MODEL_SHEET = 'model'
    return MODEL_SHEET

if __name__ == "__main__":
    arg = docopt(__doc__)
    
    file = get_abs_filepath(arg)
    sheet = get_model_sheet(arg)    
    
    make_xl_model(file, sheet)
    
    
    
