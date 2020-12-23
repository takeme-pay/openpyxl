# openpyxl
This module provides wrapper for [openpyxl](https://pypi.org/project/openpyxl/).

## Installation
```bash
pip install takeme-openpyxl
```

## Usage (Reader)
```python
import takeme_openpyxl

INPUT_FILE_NAME = 'input.xlsx'
INPUT_COLUMNS = {
    'column1': {
        'label': 'Label 1',
        'column': 'A'
    },
    'column2': {
        'label': 'Label 2',
        'column': 'B'
    }
}

reader = takeme_openpyxl.Reader(INPUT_FILE_NAME, INPUT_COLUMNS)
if reader.validate_labels() === True:
    row1_column1 = reader.get_value('column1')
    row1_column2 = reader.get_value('column2')
    print('Column1: {}, Column2: {}'.format(row1_column1, row1_column2)

    reader.next_row()

    row2_column1 = reader.get_value('column1')
    row2_column2 = reader.get_value('column2')
    print('Column1: {}, Column2: {}'.format(row1_column1, row1_column2)

reader.close()
```

## Usage (Writer)
```python

OUTPUT_FILE_NAME = 'output.xlsx'
OUTPUT_COLUMNS = {
    'column1': {
        'label': 'Label 1',
        'column': 'A'
    },
    'column2': {
        'label': 'Label 2',
        'column': 'B'
    }
}

writer = takeme_openpyxl.Writer(OUTPUT_COLUMNS)
writer.write('column1', 'Row1: Column1')
writer.write('column2', 'Row1: Column2')

writer.next_row()

writer.write('column1', 'Row2: Column1')
writer.write('column2', 'Row2: Column2')

writer.flush(OUTPUT_FILE_NAME)
```
