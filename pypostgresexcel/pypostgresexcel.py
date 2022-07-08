import string
import time
import psycopg2
import xlsxwriter


class PyPostExcel:
    def __init__(self, db_name=None, table=None, password=None, username=None, host=None, port="5432", header=True):
        self.rows = []
        self.columns = None
        self.db_name = db_name
        self.table = table
        self.password = password
        self.username = username
        self.host = host
        self.port = port
        self.db_connection = psycopg2.connect(database=self.db_name, user=self.username, password=self.password,
                                              host=self.host, port="5432")
        self.db_cursor = self.db_connection.cursor()

    def writeTable(self, file_name=f'{time.time()}', sheet_name='sheet1', table=None):
        self.getData(table=table)
        workbook = xlsxwriter.Workbook(f'{file_name}.xlsx', {'default_date_format': 'dd/mm/yyyy'})
        worksheet = workbook.add_worksheet(sheet_name)
        print(self.tableHeader())
        worksheet.add_table(f'A1:{string.ascii_uppercase[len(self.rows[0])-1]}{len(self.rows) + 1}',
                            {'data': self.rows, 'columns': self.tableHeader()})
        workbook.close()

    def getData(self, query="SELECT * FROM employee", verbose=False, table=None):

        if verbose:
            print(f'Connected to {self.host}:{self.port}')

        if table is None:
            self.db_cursor.execute(query)
        else:
            self.db_cursor.execute(f'SELECT * FROM {table}')

        rows = self.db_cursor.fetchall()
        for row in rows:
            # self.rows.append(row)
            temp = []
            for j in row:
                temp.append(j)

            self.rows.append(temp)
            if self.columns is None:
                self.columns = [[] for _ in range(len(row))]

            for i in range(len(row)):
                self.columns[i].append(row[i])

    def tableHeader(self):
        self.db_cursor.execute(
            f"SELECT * FROM information_schema.columns WHERE table_name='employee' order by ordinal_position")
        headers = []
        tempHeader = {}
        headers_schema = self.db_cursor.fetchall()
        for header in headers_schema:
            tempHeader['header'] = header[3]
            headers.append(tempHeader)
            tempHeader = {}

        return headers
