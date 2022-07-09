# PyPostExcel

PyPostExcel is a python library that export a Postgres DB to an Excel File

## Requirements

```bash
pip install psycopg2
pip install xlsxwriter
```

## Database Layout

![image](https://user-images.githubusercontent.com/101131248/178107609-5e6289e6-a728-4f2b-8866-4301c0fbbccd.png)


## Usage

```python
from pypostgresexcel import PyPostExcel


def main():
    agent = PyPostExcel(db_name='postgres', table='employee', password='demo', username='postgres', host='localhost',
                        date='date')

    agent.run(main_data_title='Personal Data', main_data=['first_name', 'id', 'last_name', 'age', 'join_date'],
              secondary_data=['supervisor_rating', 'ai_rating', 'date'])
    agent.closeWorkbook()

    print(agent.data)


if __name__ == '__main__':
    main()

```

## Arguments

| Arg | Description |
| --- | --- |
| `date` | Name of the date column in the database, that column will write the first row |
| `main_title` | title for information that comes with the root table in this case employee |
| `main_data` | columns that should be showed in the final results from the root table |
| `secondary_data` | same as main data but for the child table |

## Documantation of each function

| Function | Description |
| --- | --- |
| `def tableHeader(self, table: str):` | arg: table = name of the table to get its headers |
| `def setRowSize(self, row: int, size: float):` | change row size, row: row index, size: the size of the row |
| `def setColumnSize(self, col: str, size: float):` | same as setRowSize but for columns |
| `def InitializeFormats(self) -> None:` | initialize formats and styles for the template |
| `getTable` | get data from the DB, left join employee with employee performance |
| `def TargetedHeader(self, targetHeader, table):` | it checks if the main headers provided by the user are in the DB |
| `def getItemByName(self, name: str):` | it takes a name from the header as argument ans return an index assuming the user didn't provide the headers in the same order as the DB |
| `def CoreSection(self, data, main_data, secondary_data, supervisor=False):` | responsible of writing the core data (the records from the DB) |


