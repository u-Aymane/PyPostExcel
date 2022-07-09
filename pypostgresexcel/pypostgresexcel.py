import psycopg2
import xlsxwriter
from xlsxwriter import utility


class PyPostExcel:
    def __init__(self, db_name, table, password, username, host, date, port="5432"):
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
        self.workbook = xlsxwriter.Workbook('test.xlsx', {'default_date_format': 'dd/mm/yyyy'})
        self.worksheet = self.workbook.add_worksheet()
        self.data = {}
        self.table_root = 'employee'
        self.table_child = 'employee_performance'
        self.data_rows = []
        self.current_row = 2
        self.header_root = self.tableHeader(self.table_root)
        self.header_child = self.tableHeader(self.table_child)
        self.years = []
        self.col = 1
        self.date = date

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
                               f'LEFT JOIN {self.table_child} '
                               f'ON {self.table_root}.id = {self.table_child}.id_employee')
        rows = self.db_cursor.fetchall()
        headers = self.tableHeader(self.table_root) + self.tableHeader(self.table_child)

        for row in rows:
            self.data_rows.append(row)
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

    def getItemByName(self, name: str):
        header = self.header_root + self.header_child
        return header.index(name)

    def CoreSection(self, data, main_data, secondary_data, supervisor=False):
        if supervisor:
            self.setRootSize(self.current_row)
            self.worksheet.write(self.current_row, 0, data[self.getItemByName(main_data[0])], self.root_format)
        else:
            self.worksheet.write(self.current_row, 0, data[self.getItemByName(main_data[0])], self.child_format)

        for i in main_data[1:]:
            self.worksheet.write(self.current_row, self.col, data[self.getItemByName(i)])
            self.col += 1

        for year in self.years:
            date = data[self.getItemByName(self.date)]
            if date is not None and date.year == year:
                for i in secondary_data:
                    self.worksheet.write(self.current_row, self.col, data[self.getItemByName(i)])
                    self.col += 1

            else:
                self.col += len(secondary_data)

        self.current_row += 1
        self.col = 1

    def run(self, main_data_title: str, main_data: list, secondary_data: list):  # [id, first_name...etc] , [rating, performance, date...]
        self.getTable()
        self.InitializeFormats()

        #  Writing Years Sections
        self.worksheet.merge_range(f'B1:{utility.xl_col_to_name(len(main_data) - 1)}1', main_data_title,
                                   self.year_format)
        self.years = []
        current_index = len(main_data) - 1
        for items in self.data[self.date]:
            if items is not None and items.year not in self.years:
                self.years.append(items.year)
                self.worksheet.merge_range(
                    f'{self.ColToName(current_index + 1)}1:{self.ColToName(current_index + len(secondary_data))}1',
                    items.year, self.year_format)
                current_index += len(secondary_data)

        # Write Sub titles (headers) and values for principle section

        main_header = self.TargetedHeader(main_data, self.table_root) + self.TargetedHeader(secondary_data, self.table_child) * len(self.years)
        print(main_header)
        self.worksheet.write_row('A2', main_header, self.header_format)

        # Organize Data -> Supervisor/Employee

        for supervisor in self.data_rows:
            if supervisor[self.header_root.index('supervisor_id')] is None:
                self.CoreSection(supervisor, main_data, secondary_data, supervisor=True)
                for employee in self.data_rows:
                    if employee[self.getItemByName('supervisor_id')] == supervisor[self.getItemByName('id')]:
                        self.CoreSection(employee, main_data, secondary_data)

