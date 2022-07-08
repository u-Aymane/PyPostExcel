import time
import psycopg2
import xlsxwriter


class PyPostExcel:
    def __init__(self, db_name=None, table=None, password=None, username=None, host=None, port="5432"):
        self.rows = []
        self.db_name = db_name
        self.table = table
        self.password = password
        self.username = username
        self.host = host
        self.port = port
        self.db_connection = psycopg2.connect(database=self.db_name, user=self.username, password=self.password,
                                              host=self.host, port="5432")
        self.db_cursor = self.db_connection.cursor()
        self.year_format = None
        self.header_format = None
        self.root_format = None
        self.child_format = None
        self.worksheet = None
        self.workbook = None

    def setRowSize(self, row: int, size: float):
        self.worksheet.set_row(row, size)

    def setColumnSize(self, col: str, size: float):
        self.worksheet.set_column(col, size)

    def setRootSize(self, row: int, size=33.75):
        self.setRowSize(row, size)

    def InitializeFormats(self) -> None:
        self.year_format = self.workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'font_size': '18',
            'valign': 'center',
            'align': 'vcenter',
        })
        self.header_format = self.workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'font_size': '11',
            'valign': 'center',
            'align': 'vcenter',
        })
        self.root_format = self.workbook.add_format({
            'bold': False,
            'text_wrap': True,
            'font_size': '11',
            'valign': 'left',
            'align': 'vcenter',
        })  # -> worksheet.set_row(row, 33.75)
        self.child_format = self.workbook.add_format({
            'bold': False,
            'text_wrap': True,
            'font_size': '11',
            'valign': 'center',
            'indent': 4
        })

        self.setColumnSize('B:Z', 14.29)
        self.setColumnSize('A:A', 28.14)
        self.setRowSize(0, 50.25)
        self.setRowSize(1, 33.75)



    def run(self):
        self.workbook = xlsxwriter.Workbook('test.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        self.InitializeFormats()
        self.worksheet.merge_range('B1:D1', 'Year', self.year_format)
        self.worksheet.write('B2', 'sub_title', self.header_format)
        self.worksheet.write('A3', 'Supervisor_Section', self.root_format)
        self.setRootSize(2)

        for i in range(1, 6):
            self.worksheet.write(f'A{i + 3}', 'Employees', self.child_format)
        self.workbook.close()
