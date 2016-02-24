import numpy as np 

from make_xl_model import get_resulting_workbook_array_for_make

abs_filepath = "ref_file.xls"

expected_array_string = """[['' '2010' '2011' '2012' '2013' '2014' '2015' '2016' '2017' '2018']
 ['x' 100.0 101.0 102.5 100.8 95.5 102.5 105.0 '=I4' '=J4']
 ['y' 3360.0 2700.0 500.0 1200.0 4800.0 5280.0 6336.0 '=H3*I5' '=I3*J5']
 ['x_fut' nan nan nan nan nan nan nan 102.5 105.0]
 ['y_rog' nan nan nan nan nan nan nan 1.1 1.2]
 ['is_forecast' 0.0 0.0 0.0 0.0 0.0 0.0 0.0 1.0 1.0]]"""

assert expected_array_string == np.array_str(get_resulting_workbook_array_for_make(abs_filepath))

