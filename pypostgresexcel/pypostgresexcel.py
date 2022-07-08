import psycopg2
import xlsxwriter
from xlsxwriter import utility


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
        self.workbook = xlsxwriter.Workbook('test.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        self.data = {}
        self.table_root = 'employee'
        self.table_child = 'employee_performance'

    def tableHeader(self, table: str):
        self.db_cursor.execute(
            f"SELECT * FROM information_schema.columns WHERE table_name='{table}' order by ordinal_position")
        headers = []
        headers_schema = self.db_cursor.fetchall()
        for header in headers_schema:
            headers.append(header[3])

        return headers

    def setRowSize(self, row: int, size: float):
        self.worksheet.set_row(row, size)

    def setColumnSize(self, col: str, size: float):
        self.worksheet.set_column(col, size)

    def setRootSize(self, row: int, size=33.75):
        self.setRowSize(row, size)

    def closeWorkbook(self):
        self.workbook.close()

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

        self.setColumnSize('B:Z', 16)
        self.setColumnSize('A:A', 28.14)
        self.setRowSize(0, 50.25)
        self.setRowSize(1, 33.75)

    def getTable(self):
        self.db_cursor.execute(f'SELECT * FROM {self.table_root} '
                               f'JOIN {self.table_child} '
                               f'ON {self.table_root}.id = {self.table_child}.id_employee')
        rows = self.db_cursor.fetchall()
        headers = self.tableHeader(self.table_root) + self.tableHeader(self.table_child)

        for row in rows:
            i = 0
            for header in headers:
                if header not in self.data.keys():
                    self.data[header] = []
                self.data[header].append(row[i])
                i += 1

    def ColToName(self, val: int):
        return utility.xl_col_to_name(val)

    def TargetedHeader(self, targetHeader, table):
        header = []
        for i in targetHeader:
            if i in self.tableHeader(table):
                header.append(i)

        return header

    def OrganizeFile(self, main_data_title: str, main_data: list, secondary_data: list,
                     date='date'):  # [id, first_name...etc] , [rating, performance, date...]

        self.InitializeFormats()

        #  Writing Years Sections
        self.worksheet.merge_range(f'B1:{utility.xl_col_to_name(len(main_data) - 1)}1', main_data_title,
                                   self.year_format)
        years = []
        current_index = len(main_data) - 1
        for items in self.data[date]:
            if items is not None and items.year not in years:
                years.append(items.year)
                self.worksheet.merge_range(
                    f'{self.ColToName(current_index + 1)}1:{self.ColToName(current_index + len(secondary_data))}1',
                    items.year, self.year_format)
                current_index += len(secondary_data)

        # Write Sub titles (headers) and values for principle section

        main_header = self.TargetedHeader(main_data, self.table_root) + self.TargetedHeader(secondary_data,
                                                                                            self.table_child) * len(
            years)
        print(main_header)
        self.worksheet.write_row('A2', main_header, self.header_format)

    def run(self):
        self.InitializeFormats()
        self.worksheet.merge_range('B1:D1', 'Year', self.year_format)
        self.worksheet.write('B2', 'sub_title', self.header_format)
        self.worksheet.write('A3', 'Supervisor_Section', self.root_format)
        self.setRootSize(2)

        for i in range(1, 6):
            self.worksheet.write(f'A{i + 3}', 'Employees', self.child_format)
        self.workbook.close()
