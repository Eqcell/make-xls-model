"""
   Make Excel file with spreadsheet model based on user-defined 
   historic data, equations, control variables and spreadsheet parameters.    
   
Usage:   
   mxm.py --selftest
   mxm.py <SPEC_FILE.xls> <OUTPUT_FILE.xls>
   mxm.py --markup <YAML_FILE>
"""

"""
TO BE CORRECTED:
- no variable names written to file
- no running of all doctest tests  

LIMITATIONS (by design):
- value precisions in output xls file not corrected (can be 0 and 2)
- one sheet ber file
- decorations of dt_before_equations are hard-coded
- eats cyrulic in NAMES
"""

from docopt import docopt
from data_source import get_mock_specification
from read_xls import get_specification
from xl_fill import make_wb_array
from xls_io import write_output_to_xls

arg = docopt(__doc__)
specfile = arg["<SPEC_FILE.xls>"]


if arg['--selftest']:
    model_spec, view_spec = get_mock_specification()
elif specfile is not None:
    try:
       model_spec, view_spec = get_specification(specfile)
    except IOError:
        # todo: is this correct exception?
        raise ValueError("File not found: " + specfile)    
    except:
        # todo: is this correct exception?
        raise ValueError("Cannot read specification from file: " + specfile)
else:
    raise ValueError("No inputs provided for script.") 

# main job of creating resulting workbook array 'wb_array', dumpable to Excel
wb_array = make_wb_array(model_spec, view_spec)

# dump 'wb_array' to Excel
write_output_to_xls(wb_array, view_spec)



#----------------------------------------------------------------------------
#
#     if no user specification - write all controls in orginal order,
#                              - followed by data without controls in orginal order
#                              - year depth as in historic data
#

#
# 2015-07-20 10:46 AM
#
# Outline:
# - write to xls
# - one input xls file
# - larger xls file (flatten folder)
# - experiment with actual data - full cycle (flatten folder)
# - (source + markup) vs flattened db vs db dump
#
