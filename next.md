# PRIORITY
-----------------------
## 1 Tests:

- [ ] discover and run tests in one command
- [ ] use pytest for extended testing
- [ ] show test coverage (doctest and pytest)
- [ ] write essential tests (20-40% coverage?)
-----------------------
## 2 Module structure:

- [ ] write more docstrings
- [ ] rename config.py
- [ ] rearrange code to have fewer modules
- [+] depreciate sympy
-----------------------
## 3 Update model behaviour:
- [ ] read inputs from sheet
- [ ] write results to model sheet
- [ ] define calculable cells on model
- [ ] add is_forecast row
-----------------------
## 4 Excel 'model' sheet
- [ ] delete everything on 'model' sheet before writing to it 
- [ ] OR delete and create sheet 'model'
-----------------------
## 5 Eq preparser:
- [ ] symplify equation dicts
- [ ] merge preparser with eq parser
- [ ] only current period vars must be allowed on left side. Valid: "x(t)= x(t-1)". Not Valid: "x(t+1) = x(t)"
- [ ] add #comments to equations 
- [ ] no '=' -> not a formula
-----------------------
## 6 *ar* layout
- [ ] add more columns in *ar* to the left of label columns, use functions like 
    foo(var_label) to populate these column
- [ ] eats cyrillic var descriptions
- [ ] add empty rows before *label* in ar

# NOT TODO
## 1 Vaidation:
- [ ] may need to check for continuity of years labels
-----------------------
## 3 Documentation:
- [ ] sphinx
-----------------------
## Excel IO:
- [ ] add a button to run models 
- [ ] make Win32 exe file
-----------------------

# DONE:
## 2 Merge df:
- [+] merge df1, df2 in a right way
-----------------------
