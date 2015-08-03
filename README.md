Make an Excel spreadsheet model based on historic data, equations and control parameters.

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

## Requirements

The script is executed in [Anaconda](https://store.continuum.io/cshop/anaconda/) environment. Formal requirements.txt is to follow. 

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
