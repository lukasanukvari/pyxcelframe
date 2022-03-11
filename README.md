# PyxcelFrame

Tools for more specialized interactions between the Pandas DataFrames and the Excel worksheets.

## Install

`pip install pyxcelframe`

## Usage

### Examples

Let's suppose that we have an Excel file named **"numbers.xlsx"** with the sheet
named **"Dictionary"** in which we would like to insert the ___pandas.DataFrame___.


Import ___pandas___ and create an example ___DataFrame___ (which will be inserted into the Excel worksheet):

```python
import pandas as pd


ex = {
    'Num': [1, 2, 3, 4],
    'AfterFirstBlankCol': 'AfterFirstBlank',
    'Descr': ['One', 'Two', 'Three', 'Four'],
    'AfterSecondBlankCol': 'AfterSecondBlank.',
    'Squared': [1, 4, 9, 16],
    'Binary:': ['1', '10', '11', '100']
}

df = pd.DataFrame(ex)
```

- Import ___openpyxl.load_workbook___ and open **numbers.xlsx** - Our Excel workbook;
- Get - **Dictionary** our desired sheet:

```python
from openpyxl import load_workbook


workbook = load_workbook('numbers.xlsx')
worksheet = workbook['Dictionary']
```
#### Functions

##### 1. `column_last_row(worksheet, column_name)`

- If we had to get the last non-empty row in coolumn __A__ of Excel worksheet called __Dictionary__ and if we definitely knew that there would not be more than __10000__ row records in that column:


_NOTE: By default `count_from` will be __1048576__, because that number is the total amount of the rows in an Excel worksheet._
```python
from pyxcelframe import column_last_row


column_last_row(worksheet=worksheet, column_name=['A'], count_from=10000)
```

##### 2. `copy_cell_style(cell_src, cell_dst)`

- Let's say, we have a cell in Excel __Dictionary__ worksheet that we would like to copy the style from,
and it is __O3__;
- Let __O4__ be our destination cell:

_NOTE: If we wanted to copy that style to more than one cell, we would simply use the loop
depending on the locations of the destination cells._

```python
from pyxcelframe import copy_cell_style


copy_cell_style(cell_src=worksheet['O3'], cell_dst=worksheet['O4'])
```

##### 3. `sheet_to_sheet(filename_sheetname_src, filename_sheetname_dst, calculated)`

- Let's say that we have two Excel files, and we need specific sheet from one file
to be completely copied to another file's specific sheet;
- `filename_sheetname_src` is the parameter for one file -> sheet the data
to be copied from ___(tuple(['FILENAME_SRC', 'SHEETNAME_SRC']))___;
- `worksheet_dst` is the parameter for the destination ___Worksheet___ the data
to be copied to ___(openpyxl.worksheet.worksheet.Worksheet)___;
- Let's assume that we have __file_src.xlsx__ as src file and for `worksheet_src` we can
use its __CopyThisSheet__ sheet.
- We can use __output.xlsx__ -> __CopyToThisSheet__ sheet as the destination worksheet, for which
we already declared the ___Workbook___ object above.

_NOTE: We are assuming that we need all the formulas (where available) from the source sheet, not calculated data, so we set `calculated` parameter to __False__._

```python
from pyxcelframe import sheet_to_sheet


worksheet_to = workbook['CopyToThisSheet']

sheet_to_sheet(filename_sheetname_src=('file_src.xlsx', 'CopyThisSheet'),
               worksheet_dst=worksheet_to,
               calculated=False)
```

##### 4. `insert_frame(worksheet, dataframe, col_range, row_range, num_str_cols, skip_cols, headers)`

- From our package ___pyxcelframe___ import function ___insert_frame___;
- Insert `ex` - ___DataFrame___ into our sheet twice - with and without conditions:

```python
from pyxcelframe import insert_frame


# 1 - Simple insertion
insert_frame(worksheet=worksheet, dataframe=df)

# 2 - Insertion with some conditions
insert_frame(worksheet=worksheet,
             dataframe=df,
             col_range=(3, 0),
             row_range=(6, 8),
             num_str_cols=['I'],
             skip_cols=['D', 'F'],
             headers=True)
```

In the first insertion, we did not give our function any arguments, which means the ___DataFrame___
`ex` will be inserted into the __Dictionary__ sheet in the area __A1:F4__ (without the headers).

However, with the second insertion we define some conditions:

- `col_range=(3, 0)` - This means that insertion will be started at the Excel column with the
index 3 (column __C__) and will not be stopped until the very end, since we gave 0 as the
second element of the tuple

- `row_range=(6, 8)` - Only in between these rows (in Excel) will the ___DataFrame___ data be inserted,
which means that only the first row (since the `headers` is set to _True_) from `ex` will be inserted into the sheet

- `num_str_cols=['F']` - Another condition here is to not convert _Binary_ column values to int.
If we count, this column will be inserted in the Excel column __F__, so we tell the function to leave
the values in it as string

- `skip_cols=['D', 'F']` - __D__ and __F__ columns in Excel will be skipped and since our worksheet
was blank in the beginning, these columns will be blank (that is why I named the columns in the
___DataFrame___ related names)

- `headers=True` - This time, the ___DataFrame___ columns will be inserted, too, so the overall
insertion area would be __C6:J8__

##### 5. `insert_columns(worksheet, dataframe, columns_dict, row_range, num_str_cols, headers)`

- From our package ___pyxcelframe___ import function ___insert_columns___;
- Insert `ex` - ___DataFrame___ into our sheet according to the `cols_dict` - ___Dict___ which contains the `ex` ___DataFrame's___ column names as the keys and the `worksheet` Excel ___Worksheet's___ column names as the values:

_NOTE: Only those columns that are included as the `cols_dict` keys will be inserted into the `worksheet` from the `ex` ___DataFrame___; Also, all the other parameters are similar to the parameters of the `insert_frame` function, so we will only be giving the required arguments for this example._

```python
from pyxcelframe import insert_columns


# Column "Num" of the `ex` DataFrame will be
# inserted to the "I" column of the `worksheet`
# "Descr" to "J"
# "Squared" to "L"
cols_dict = {
    "Num": "I",
    "Descr": "J",
    "Squared": "L"
}

insert_columns(worksheet=worksheet,
               dataframe=df,
               columns_dict=cols_dict)
```

- Finally, let's save our changes to a new Excel file:

```python
workbook.save('output.xlsx')
```

###### For the really detailed description of the parameters, please see `__doc__` attribute of the above functions.

#### Full Code

```python
import pandas as pd
from openpyxl import load_workbook
from pyxcelframe import copy_cell_style, \
                        insert_frame, \
                        insert_columns, \
                        sheet_to_sheet, \
                        column_last_row


ex = {
    'Num': [1, 2, 3, 4],
    'AfterFirstBlankCol': 'AfterFirstBlank',
    'Descr': ['One', 'Two', 'Three', 'Four'],
    'AfterSecondBlankCol': 'AfterSecondBlank.',
    'Squared': [1, 4, 9, 16],
    'Binary:': ['1', '10', '11', '100']
}

df = pd.DataFrame(ex)

workbook = load_workbook('numbers.xlsx')
worksheet = workbook['Dictionary']

# Column "Num" of the `ex` DataFrame will be
# inserted to the "I" column of the `worksheet`
# "Descr" to "J"
# "Squared" to "L"
cols_dict = {
    "Num": "I",
    "Descr": "J",
    "Squared": "L"
}


# Get the last non-empty row of the specific column
column_last_row(worksheet=worksheet, column_name=['A'], count_from=10000)


# Copy the cell style
copy_cell_style(cell_src=worksheet['O3'], cell_dst=worksheet['O4'])


# Copy the entire sheet
worksheet_to = workbook['CopyToThisSheet']

sheet_to_sheet(filename_sheetname_src=('file_src.xlsx', 'CopyThisSheet'),
               worksheet_dst=worksheet_to,
               calculated=False)


# Insert DataFrame into the sheet

## 1 - Simple insertion
insert_frame(worksheet=worksheet, dataframe=df)

## 2 - Insertion with some conditions
insert_frame(worksheet=worksheet,
             dataframe=df,
             col_range=(3, 0),
             row_range=(6, 8),
             num_str_cols=['I'],
             skip_cols=['D', 'F'],
             headers=True)

## 3 - Insertion according to the `cols_dict` dictionary
insert_columns(worksheet=worksheet,
               dataframe=df,
               columns_dict=cols_dict)


workbook.save('output.xlsx')
```
