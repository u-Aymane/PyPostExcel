from pypostgresexcel import PyPostExcel


def main():
    agent = PyPostExcel(db_name='postgres', table='employee', password='demo', username='postgres', host='localhost')
    # agent.writeXLSX(table='employee', file_name='table_1')
    agent.writeTable(table='employee', file_name='table_1')


if __name__ == '__main__':
    main()
