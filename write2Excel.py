import xlsxwriter   # xlsxwriter module

class write2Excel():
    """ This class uses the xlsxwriter module to write the content of a list of tuples to an Excel file
        :filename => name of the Excel file to be written
        :ws_name => name of the worksheet to be written to
    """

    def __init__(self, filename = "myFile.xlsx", ws_name = 'myWorkSheet'):
        self.filename = filename    # name of the excel file to be writtent to
        self.ws = ws_name           # name of the worksheet to be written to
        self.contents_list = [("Header 1", "Header 2", "Header 3"),
            ("row1-col1","row1-col2","row1-col3"),
            ("row2-col1","row2-col2","row2-col3")
        ]

    def formatted_write(self, row, col, content, wb, ws, fmt):
        """ Writes cells with a particular format. """
        if fmt == "bold":
            format = wb.add_format({"bold": True})
        elif fmt == "currency":
            format = wb.add_format({'num_format': '$#,##0'})
        ws.write(row, col, content, format)

    def write_list_contents(self):
        """ Writes the content of a list of tuples into an Excel file"""
        # creating a workbook
        workbook = xlsxwriter.Workbook(self.filename)
        # adding a worksheet
        worksheet = workbook.add_worksheet(self.ws)
        row = 0
        col = 0
        content = ""
        # writes the content of every column in every row in the contents list
        for cols in self.contents_list:
            for col_content in cols:
                if row == 0:    # writes the header row in bold
                    format = "bold"
                    self.formatted_write(row, col, col_content, workbook, worksheet, format)
                else:
                    worksheet.write(row, col, col_content)
                col += 1
            col = 0
            row += 1
        # closing the workboook, now that we are done
        workbook.close()

if __name__ == "__main__":
    x = write2Excel()
    x.write_list_contents() # writes the contents of a fake list into a dummy file