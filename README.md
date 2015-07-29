Make an Excel spreadsheet model based on historic data, equations and control parameters.

- historic data, equations and control parameters are listed on individual sheets of input Excel file (by default - 'data', 'equations' and 'controls')
- model is written to 'model' sheet of Excel file 

The script intends to:
- separate historic data from model/forecast specification 
- explictly show all forecast parameters (usually hidden inside a spreadsheet)
- explicitly show system of equations governing the forecast 

The script does not intend to:
- do forecast calculations anywhere except Excel/OpenOffice
- resolve/optimise formulas, including circular references
- spread Excel model to many sheets

## Requirements

The script is executed in [Anaconda](https://store.continuum.io/cshop/anaconda/) environment. Formal requirements.txt is to follow. 

## Interface
```python mxm.py <xlfile> [--make | --update]``` will make a new speadsheet model or update existing one.   

## Trial runs
There are two files with simple models included: spec.xls and spec2.xls. One can see the results of creating a model by runnin the following:
```
python mxm.py spec.xls --make
python mxm.py spec2.xls --make
```
or
```python make_xl_model.py```

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
