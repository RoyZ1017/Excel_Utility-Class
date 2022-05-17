import win32com.client


class ExcelUtility:
    row_count = 0

    def __init__(self, excel_file_name, sheet_name):
        self.sheet_name = sheet_name
        self.excel_file_name = excel_file_name

    def open_excel(self):
        try:
            excel = win32com.client.Dispatch('Excel.Application')
            excel.Visible = False
            wb = excel.Workbooks.Open(self.excel_file_name)
            return [excel, wb]
        except Exception as e:
            print(e)
            return None

    def cell_value(self, wb, row, column):
        try:
            all_data = wb.Worksheets(self.sheet_name).UsedRange
            read_cell = str(all_data.Cells(row, column))
            return read_cell
        except Exception as e:
            print(e)
            return None

    def num_row(self, wb):
        try:
            all_data = wb.Worksheets(self.sheet_name).UsedRange
            get_rows = all_data.Rows.Count
            self.row_count = get_rows
            return get_rows
        except Exception as e:
            print(e)
            return None

    def num_column(self, wb):
        try:
            all_data = wb.Worksheets(self.sheet_name).UsedRange
            get_columns = all_data.Columns.Count
            return get_columns
        except Exception as e:
            print(e)
            return None

    def write_in_cell(self, row, column, message, wb):
        try:
            write_data = wb.Worksheets(self.sheet_name)
            write_data.Cells(row, column).Value = message
            wb.Save()
        except Exception as e:
            print(e)
            return None

    def delete_cell(self, row, column, wb):
        try:
            write_data = wb.Worksheets(self.sheet_name)
            write_data.Cells(row, column).Value = None
            wb.Save()
        except Exception as e:
            print(e)
            return None

    def two_d_array(self, row, column, wb):
        try:
            array = []
            all_data = wb.Worksheets(self.sheet_name).UsedRange
            for r in range(0, row):
                array.append([])
                for c in range(0, column):
                    read_cell = str(all_data.Cells(r + 2, c + 1))
                    array[r].append(read_cell)
            return array
        except Exception as e:
            print(e)
            return None

    def close(self, wb, excel):
        try:
            wb.Close()
            excel.Application.Quit()
        except Exception as e:
            print(e)
            return None



