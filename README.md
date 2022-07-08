# BC Skills - Projects

PyPostExcel is a python library that export a Postgres DB to an Excel File

## Requirements

```bash
pip install psycopg2
pip install xlsxwriter
```

## Usage

```python
from pypostgresexcel import PyPostExcel


def main():
    agent = PyPostExcel(db_name='postgres', table='employee', password='demo',
                        username='postgres', host='localhost')
    agent.writeXLSX(table='employee', file_name='table_1')


if __name__ == '__main__':
    main()

```
