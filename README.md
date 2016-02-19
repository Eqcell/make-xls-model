# Scope of work

Core functionality (engine): autogenerate formulas in Excel cells based on variable names and list of equations. 
Final use (application): make clean Excel spreadsheet model with formulas based on historic data, equations and control parameters.

# Main idea

Spreadsheet models in Excel often become a mess: 
- the whole picture what equations cannot be seen easily
- cannot guarantee it is the same equation across all cells in row/column 
- cannot replicate or amend many formulas in Excel file fast
- control parameters may be hidden somewhere and it is unclear what really governs your forecast.   

By 'spreadsheet models' I mean a simple structure with some historic values for time series, some assigned future values for control parameters and equations that tell you how control parameters affect the rest of the variables. 

## Minimal example

GDP forecast value is a function of prvious yeat value, deflator (Ip) and real growth rate (Iq):

|   | A   | B     | C     |
|---|-----|-------|-------|
| 1 | GDP | 23500 | 25415 |
| 2 | Iq  |       | 1,05  |
| 3 | Ip  |       | 1,03  |

In C1 we have a formula ```=B1*C2*C3```. What I want to have is be able to generate this formula from a string in cell A4 below.

|   | A   | B     | C     |
|---|-----|-------|-------|
| 1 | GDP | 23500 |       |
| 2 | Iq  |       | 1,05  |
| 3 | Ip  |       | 1,03  |
| 4 | GDP = GDP[t]*Iq*It  |       | 1,03  |

#File specification and script behaviour

- historic data, equations and control parameters are listed on individual sheets of input Excel file (by default - 'data', 'equations' and 'controls')
- model is written to 'model' sheet of Excel file 
- 'model' sheet can be generated from 'data', 'equations' and 'controls' sheets (```-M``` key)
- once 'model' sheet is created one can change data, controls and equation solely in it and refresh formulas in cells (```-U``` key)

The script intends to:
- separate historic data from model/forecast specification 
- explictly show all forecast parameters 
- explicitly show equations in the model  
- make a stand-alone Excel fiel with no dependecies or VBA code, just new clean formulas in it.

The script does not intend to:
- do any forecast calculations outside Excel/OpenOffice
- resolve/optimise formulas, including circular references
- spread Excel model to many sheets

## Requirements

The script is executed in [Anaconda](https://store.continuum.io/cshop/anaconda/) environment, we use Python 3.5.

Formal requirements.txt is to follow. 

## Interface
```python mxm.py <xlfile> [-M | -U]```     
\-M will overwrite sheet 'model' with a new sheet derived from sheets 'data', 'controls', 'equations' and 'names'  
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

#Other repos

parts of the code may be found in my other repos
- <https://github.com/epogrebnyak/eqcell>
- <https://github.com/epogrebnyak/roll-forward> (private)

