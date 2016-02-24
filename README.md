# Main idea explained

In financial analysis and economic forecasting there is a common type of 'spreadsheet models' in Excel which include the following:
- there is some observed historic data for time series (e.g. balance sheet items); 
- forecast is made by assigning future values to some control parameters (growth rates, elasticities, ratios, etc);
- there are that equations link control parameters to the rest of the variables. 

Large Excel files of this kind often become a mess: 
- the whole picture of equations cannot be seen easily
- cannot guarantee it is the same equation across all cells in row/column 
- cannot replicate or amend many formulas in Excel file fast
- control parameters may be hidden somewhere and it is unclear what really governs your forecast.

This problem grows bigger with your file size, model complexity and number of people working on it. However, we are binded to use  Excel because it has a great user interface, people can experiment with their own changes quickly, can share a model as one file with no extra dependencies.  

```make-xls-model``` is a tool where I provide historic data, control parameter values and a list of equations on separate sheets 
in Excel file and generate resulting spreadsheet model on another sheet in this file. 

The resulting file should look the same as if I worked in Excel only - no extra dependecies or VBA code, just a regular stand-alone Excel file with proper formulas in cells.

With spreadsheet models of about 20-50 or more equations I assume there should be a big productivity gain, espacially if model structure is sometimes reviewed. 

## Minimal example

GDP forecast value is a function of previous yeat value, deflator (Ip) and real growth rate (Iq):

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
| 4 | GDP = GDP[t-1]\*Iq\*It  |       |  |

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

## Interface
```python mxm.py <xlfile> [-M | -U]```    

- ```-M``` will overwrite sheet 'model' with a new sheet derived from sheets 'data', 'controls', 'equations' and 'names'  
- ```-U``` will only update formulas on sheet 'model'   

## Examples 

There are several Excel files provided in [examples](examples) folder. To invoke ```make-xls-model``` you may use [examples.bat](examples/examples.bat). 

## Assumptions and limitations

- uses ```xlwings``` and runs on Windows only, no linux
- annual labels only (continious integers)
- data organised by row only
- one model sheet in file
- a variable appears only once on model sheet

#Other repos

Parts of the code may be found in my other repos
- <https://github.com/epogrebnyak/eqcell>
- <https://github.com/epogrebnyak/roll-forward> (private)

# Scope of work

Autogenerate formulas in Excel cells based on variable names and list of equations (core functionality/engine)

Make clean Excel spreadsheet model with formulas based on historic data, equations and control parameters (final use / application) 

## Requirements

The script is executed in [Anaconda](https://store.continuum.io/cshop/anaconda/) environment, we use Python 3.5.

Formal requirements.txt is to follow.
