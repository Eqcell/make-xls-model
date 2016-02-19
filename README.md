# Main idea

Based on my expereice in economic spreadsheet modelling Excel files often become a mess: 
- you do not see the whole picture what equations were used
- cannot guarantee it is the same equation across all cells in row/column 
- cannot replicate or amend many formulas Excel file fast
- your control parameters may be hidden somewhere and it is unclear what really governs your forecast.   

By saying 'model' or 'forecast' I mean a simple structure where there are some historic values for time series, some known assigned future values for control parameters and equations that tell you how controls affest the rest of the variables. 

In a minamal example you can have GDP forecast value to be a function of deflator (Ip) and real growth rate (Iq). An Excel sheet will look like below:  

--|-----|-------|-----
1 | GDP | 23500 | 25415
2 | Iq  |       |  1,05
3 | Ip  |       |  1,03



```
     2010   2011  
GDP   500   
Ip          1.05
Iq          


```

So I wanted a tool  

# Scope 
Core functionality (engine): autogenerate formulas in Excel cells based on variable names and list of equations. 
Final use (application): make clean Excel spreadsheet model with formulas based on historic data, equations and control parameters.

Scope:


Features:
- historic data, equations and control parameters are listed on individual sheets of input Excel file (by default - 'data', 'equations' and 'controls')
- model is written to 'model' sheet of Excel file 
- 'model' sheet can be updated without modifing 'data', 'equations' and 'controls' sheets

The script intends to:
- separate historic data from model/forecast specification 
- explictly show all forecast parameters (usually hidden inside spreadsheet rows)
- explicitly show equations that link previous period variables to current period and produce the forecast 

The script does not intend to:
- do forecast calculations outside Excel/OpenOffice
- resolve/optimise formulas, including circular references
- spread Excel model to many sheets

NOTE: parts of the code may be found in my other repos
- <https://github.com/epogrebnyak/eqcell>
- <https://github.com/epogrebnyak/roll-forward> (private)


## Requirements

The script is executed in [Anaconda](https://store.continuum.io/cshop/anaconda/) environment, we use Python 3.5.

Formal requirements.txt is to follow. 

## Interface
```python mxm.py <xlfile> [-M | -U]```     
\-M will overwrite sheet 'model' with a new one derived from sheets 'data', 'controls', 'equations' and 'names'  
\-U will only update formulas on sheet 'model'   

## Trial runs
There are several files with simple models included (eg spec.xls and spec2.xls). One can see the results of creating or updating a model by running the following:
```
python mxm.py spec.xls -M
python mxm.py spec2.xls -M
```
or
```python make_xl_model.py```

## Assumptions and limitations

- annual labels only (continious integers)
- by row only
- one model sheet in file
- variable appears only once on model sheet

## Test coverage
```
Name               Stmts   Miss  Cover
--------------------------------------
formula_parser        93     18    81%
globals                2      0   100%
iterate_in_array      25      2    92%
make_xl_model        121     21    83%
--------------------------------------
TOTAL                241     41    83%
```
