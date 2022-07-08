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
        if header:
            self.rows.append(self.tableHeader())

    def getData(self, query="SELECT * FROM employee", verbose=False, table=None):

        if verbose:
            print(f'Connected to {self.host}:{self.port}')  # by setting verbose to true the end user can see what's happening

        if table is None:
            self.db_cursor.execute(query)  # custom query
        else:
            self.db_cursor.execute(f'SELECT * FROM {table}')  # getting all data from a table

        rows = self.db_cursor.fetchall()
        for row in rows:
            self.rows.append(row)
            if self.columns is None:
                self.columns = [[] for _ in range(len(row))]  # create arrays for each column

            for i in range(len(row)):
                self.columns[i].append(row[i])  # append data to columns arrays

    def tableHeader(self):  # get table header
        self.db_cursor.execute(
            f"SELECT * FROM information_schema.columns WHERE table_name='employee' order by ordinal_position")
        headers = []
        headers_schema = self.db_cursor.fetchall()
        for header in headers_schema:
            headers.append(header[3])

        return headers

    def writeXLSX(self, file_name=f'{time.time()}', sheet_name='sheet1', table=None):  # write the XLSX file
        self.getData(table=table)
        workbook = xlsxwriter.Workbook(f'{file_name}.xlsx', {'default_date_format': 'dd/mm/yyyy'})
        worksheet = workbook.add_worksheet(sheet_name)
        row = 0
        for rows in self.rows:
            col = 0
            for val in rows:
                worksheet.write(row, col, val)
                col += 1
            row += 1

        workbook.close()
