# Python - Excel

## Terminology
- Workbook: A Workbook (Spreadsheet) is the main file (excel file) you are creating or working with.
- Sheet: A Sheet (Worksheet) is used to split different kinds of content within the same workbook. A workbook can have one or more sheets.
- Column: A Column is a vertical line, and it is represented by an uppercase letter: A, B, C, etc..
- Row: A Row is a horizontal line, and it is represented by a number: 1, 2, 3, etc..
- Cell: A Cell is a combination of Column and Row, represented by both an uppercase letter and a number: A1, B3, C5, etc..

## Getting Started 
Install the ```openpyxl``` module.
```console
pip install openpyxl
```

Create a simple workbook
```python
from openpyxl import Workbook

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "First Name"
sheet["B1"] = "Last Name"
sheet["A2"] = "Gary"
sheet["B2"] = "Wee"

workbook.save(filename="name.xlsx")
workbook.close()
```
![Screenshot_01](https://user-images.githubusercontent.com/51909547/182057199-70cfd4a8-f966-4fbc-9bbe-395c70a6383c.png)



## Reading Data From Excel

```python
from openpyxl import load_workbook
file_name = "name.xlsx"
workbook = Workbook(file_name)
sheeet = wb.active

cell_a1 = sheet['A1']
cell_a1_value = cell_a1.value
print(f'Cell value in A1: {cell_a1_value}')

workbook.close()
```

## Manipulation of Worksheet

# Create new worksheet
```python
from openpyxl import Workbook
file_name = "sheet_manipulation.xlsx"
wb = Workbook()
ws1 = wb.create_sheet('sheet_1')    # create sheet at the last position
ws2 = wb.create_sheet('sheet_2',0)  # create sheet at the 0th position
ws3 = wb.create_sheet('sheet_3',-1) # create sheet at the second last position
wb.save(file_name)
wb.close()
```

# Change worksheet's name
```python
ws = wb['sheet_2']         # select worksheet named 'sheet_2'
ws.title = 'new_sheet_2'   
```
or
```python
ws = wb.worksheets[0]      # select the 0th worksheet 
ws.title = 'new_sheet_0'
```


To Be Continued...

Reference: https://realpython.com/openpyxl-excel-spreadsheets-python/
