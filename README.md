Excel files often [get messy](problem.md). 

```make-xls-model``` is a tool to create a spreadsheet model based on provided historic data, equations and control parameters. 
The resulting sheet contains a model with proper formulas in cells and no other dependencies.

In breif, we intend to:
- separate historic data from model/forecast specification 
- explictly show all forecast parameters 
- explicitly show all equations in the model  
- make a stand-alone Excel file with no dependencies or VBA code, just new clean formulas in it
- try make spreadsheet models replicable 

The script does not intend to:
- do any forecast calculations outside Excel/OpenOffice
- resolve/optimise formulas, including circular references
- spread Excel model to many sheets

## Simple illustration

GDP forecast value is a function of previous yeat value, deflator (Ip) and real growth rate (Iq):

|   | A   | B     | C     |
|---|-----|-------|-------|
| 1 | GDP | 23500 | 25415 |
| 2 | Iq  |       | 1,05  |
| 3 | Ip  |       | 1,03  |

In ```C1``` we have a formula ```=B1*C2*C3```.  ```make-xls-model``` can generate this formula and place in ```C1``` from a string in cell ```A4``` below.

|   | A   | B     | C     |
|---|-----|-------|-------|
| 1 | GDP | 23500 |       |
| 2 | Iq  |       | 1,05  |
| 3 | Ip  |       | 1,03  |
| 4 | GDP = GDP[t-1]\*Iq\*It  |       |  |

##Workflow:
- historic data, equations and control parameters are listed on individual sheets of Excel file (by default - 'data', 'equations' and 'controls')
- spreadsheet model is placed to 'model' sheet of Excel file
- 'model' sheet can be generated from 'data', 'equations' and 'controls' sheets (```-M``` key)
- once 'model' sheet is created one can change data, controls and equation solely on this sheet and refresh formulas in cells with  ```-U``` key

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

##One-liners

Core functionality (engine):  
Autogenerate formulas in Excel cells based on variable names and list of 

Final use (application):
Make clean Excel spreadsheet model with formulas based on historic data, equations and control parameters

##Other repos

Parts of the code may be found in my other repos
- <https://github.com/epogrebnyak/eqcell>
- <https://github.com/epogrebnyak/roll-forward> (private)

## Requirements

The script is executed in [Anaconda](https://store.continuum.io/cshop/anaconda/) environment, we use Python 3.5. Everything runs only on Windows. 

Formal requirements.txt is to follow.
