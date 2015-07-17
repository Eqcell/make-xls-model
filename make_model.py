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

# ****************************************** 
# Proxies to speed up development / testing
# *****************************************

import io
import csv

       
def get_data(label, year):
    data_proxy = [ ("GDP", 2013, 66190.11992)
        , ("GDP",    2014, 71406.3992)
        , ("GDP_IQ", 2013, 101.3407976)
        , ("GDP_IQ", 2014, 100.6404858)
        , ("GDP_IP", 2013, 105.0467483)
        , ("GDP_IP", 2014, 107.1941886)
        ]
    slice = [x for x in data_proxy if x[0] == label]
    return  [x for x in slice if x[1] == year][0]


def yield_proxy():
    names_csv_proxy ="""text    label   group   level   precision
ВВП	gdp	Нацсчета	1	0
Индекс физ.объема ВВП	gdp_Iq	Нацсчета	2	1
Дефлятор ВВП	gdp_Ip	Нацсчета	2	1"""
    
    text_stream = io.StringIO(names_csv_proxy)
    reader = csv.DictReader(text_stream, delimiter='\t')
    for row in reader:
        print (row)

yield_proxy()      
             
controls_proxy = [("GDP_IP", 2015, 95,0)
        , ("GDP_IQ", 2015, 115,0)
        ]

equations_1 = ["GDP(t) = GDP(t-1) * GDP_IP(t) / 100 * GDP_IQ(t) / 100"]

equations = equations_1

format_1 = { 'filename': "macro.xls"
        , 'sheet': "gdp_forecast"
        , 'start_year': 2014
        , 'rows': ["GDP", "GDP_IP", "GDP_IQ"]
        }

format = format_1        
  
final_dataframe_proxy = """		2014	2015	2016
ВВП	gdp	71406	=D3*E4/100*E5/100	=E3*F4/100*F5/100
Индекс физ.объема ВВП	gdp_Iq	100,6	95	102
Дефлятор ВВП	gdp_Ip	107,2	110	112
"""        
        
        
#***********************************************        
YEAR_ROW = 2 
LABEL_COLUMN = 3

label_postions = []
for j, label in enumerate(format['rows']):
    label_postions.append((label, YEAR_ROW + j - 1, LABEL_COLUMN - 1))
print (label_postions) 




