"""Make Excel file with spreadsheet model based on historic data, equations, 
   control variables and spreadsheet parameters.      
   
Usage:   
   mxm.py --selftest
   mxm.py <SPEC_FILE> 
   mxm.py --markup <YAML_FILE>
"""

from docopt import docopt
from read_xls import get_specification_from_arg
from xl_fill import make_wb_array
from xls_io import write_output_to_xls

from xlwings import Workbook, Range, Sheet
import numpy as np


def write_output_to_xls(wb_array, view_spec)
    file = view_spec['file']
    sheet = view_spec['model']
    wb = Workbook(file)
    Range(sheet, 'A1').value = ar
    Sheet(sheet).activate()
    wb.save()
    # do not close file
    # wb.close()

if __name__ = "__main__":
    arg = docopt(__doc__)

    # init model parameters
    # model_spec, view_spec = get_specification_from_arg(arg)

    # main job of creating resulting workbook array 'wb_array', dumpable to Excel
    # wb_array = make_wb_array(model_spec, view_spec)

    from data_source import _sample_for_xfill_array_after_equations 
    ar =  _sample_for_xfill_array_after_equations()

    view_spec = {'file': 'D:/make-xls-model-master/spec.xls', 
                'sheet': 'model'}

    # dump 'wb_array' to Excel
    write_output_to_xls(wb_array, view_spec)