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
