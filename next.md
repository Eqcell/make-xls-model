-[ ] create/update model sheet
-[ ] merge prparser to formula parser
-[ ] only current period vars must be allowed on left side. Valid: "x(t)= x(t-1)". Not Valid: "x(t+1) = x(t)"

Excel IO:
- [ ] add a buttin to run models 

Module structure:
-[ ] write docstngs
-[ ] rearrange code to have fewer modules
-[ ] merge preparser with eq parser
-[ ] depreciate sympy
-[ ] symplify equation dicts

Tests:
-[ ] discover and run tests in one command
-[ ] use pytest for extended testing
-[ ] show test coverage (doctest and pytest)
-[ ] write essential tests (20-40% coverage?) 

*ar* layout
-[ ] add more columns in *ar* to the left of label columns, use functions like 
    foo(var_label) to populate these column
-[ ] eats cyrillic var descriptions
-[ ] add empty rows before *label* in ar

More:
-[ ] delete everything on 'model' sheet before writing to it OR delete and 
     create sheet 'model'
-[ ] may need to check for continuity of years labels

Equations:
-[ ] add #comments to equations 
-[ ] no '=' -> not a formula

Merge df:
-[ ] merge df1, df2 in a right way

Update 'model':
-[ ] read inputs from model
-[ ] add is_forecast row

LIMITATIONS (by design):
- annual only
- byRow only 
- no var descriptions in file (pivot column needed) - NAMES
- one model sheet per file
- one occurrence of variable in reulting sheet (variable appearing only once)
- no formatting in Excel file

 