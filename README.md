# PyCellFrame

Insert Pandas DataFrames into the Excel sheet with a bunch of conditions

## Install

`pip install pycellframe`

## Usage

### Examples

Let's suppose that we have an Excel file named **"numbers.xlsx"** with the sheet
named **"Dictionary"** in which we would like to insert the ___pandas.DataFrame___.


Import ___pandas___ and create an example ___DataFrame___ (which will be inserted into the Excel sheet):

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

- Let's say, we have a cell in Excel __Dictionary__ sheet that we would like to copy the style from,
and it is __O3__;
- Let __O4__ be our destination cell:

_NOTE: If we wanted to copy that style to more than one cell, we would simply use the loop
depending on the locations of the destination cells._

```python
from pycellframe import incell_style


incell_style(cell_src=worksheet['O3'], cell_dst=worksheet['O4'])
```

- From our package ___pycellframe___ import function ___incell_frame___;
- Insert `ex` - ___DataFrame___ into our sheet twice - with and without conditions:

```python
from pycellframe import incell_frame


# 1 - Simple insertion
incell_frame(worksheet=worksheet, dataframe=df)

# 2 - Insertion with some conditions
incell_frame(worksheet=worksheet,
             dataframe=df,
             col_range=(3, 0),
             row_range=(2, 4),
             num_str_cols=('I'),
             skip_cols=('D', 'F'),
             headers=True)
```

In the first insertion, we did not give our function any arguments, which means the ___DataFrame___
`ex` will be inserted into the __Dictionary__ sheet in the area __A1:F4__ (without the headers).

However, with the second insertion we define some conditions:

- `col_range=(3, 0)` - This means that insertion will be started at the Excel column with the
index 3 (column __C__) and will not be stopped until the very end, since we gave 0 as the
second element of the tuple

- `row_range=(2, 4)` - Only in between these rows (in Excel) will the ___DataFrame___ data be inserted,
which means that only the first row (since the `headers` is set to _True_) from `ex` will be inserted into the sheet

- `num_str_cols=('F')` - Another condition here is to not convert _Binary_ column values to int.
If we count, this column will be inserted in the Excel column __F__, so we tell the function to leave
the values in it as string

- `skip_cols=('D', 'F')` - __D__ and __F__ columns in Excel will be skipped and since our worksheet
was blank in the beginning, these columns will be blank (that is why I named the columns in the
___DataFrame___ related names)

- `headers=True` - This time, the ___DataFrame___ columns will be inserted, too, so the overall
insertion area would be __C2:J4__

###### For really detailed description of the parameters, please see `incell_frame.__docs__`

- Finally, let's save our changes to a new Excel file:

```python
workbook.save('output.xlsx')
```

#### Full Code

```python
import pandas as pd
from openpyxl import load_workbook
from pycellframe import incell_style, incell_frame


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

incell_style(cell_src=worksheet['O3'], cell_dst=worksheet['O4'])

# 1 - Simple insertion
incell_frame(worksheet=worksheet, dataframe=df)

# 2 - Insertion with some conditions
incell_frame(worksheet=worksheet,
             dataframe=df,
             col_range=(3, 0),
             row_range=(2, 4),
             num_str_cols=('I'),
             skip_cols=('D', 'F'),
             headers=True)

workbook.save('output.xlsx')
```
