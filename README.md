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

## Requirements

The script is executed in [Anaconda](https://store.continuum.io/cshop/anaconda/) environment. Formal requirements.txt is to follow. 

## Interface
```   
Usage:   
   mxm.py <xlfile> [--make | --update]
```

Reference call:
```python make_xl_model.py```
or
```
python mxm.py spec.xls
python mxm.py spec2.xls
```


