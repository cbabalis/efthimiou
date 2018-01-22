""" This module reads several files (xls format),
extracts only the useful information and
writes it to a .csv file.

In particular, two scenarios are being examined. Given a file of GPS points:
    - extract all data relevant to specific routes (i.e. Attiki Odos +/-1 stop)
    - extract all routes a truck has been done.

    To run this module as an independent script, type:
    > python extract_routes.py ~/my_folder output.csv
"""


import pdb
import xlrd
import xlwt
import os
import sys
reload(sys)
sys.setdefaultencoding('utf8')

class Extraction:
    """ Extracts data from several files to a new one.
    """

    def __init__(self, src_folder, dest_file, value):
        self.file_contents = []
        self.src_folder = src_folder
        self.dest_file = dest_file
        # read the file
        self.read_folder(src_folder)
        # process data and write data to a new file
        self.process_file(value)

    def read_folder(self, folder=""):
        """ This method reads a folder which contains files as input.

        :param str folder: is the folder where the function looks
        for files recursively
        """
        # search recursively the system
        for root, dirs, files in os.walk(folder):
            for f in files:
                if f.endswith(".xls"):
                    # read the xls file
                    xls_file = os.path.join(root, f)
                    self._read_xls(xls_file)

    def _read_xls(self, xls_file):
        """ This method reads a single xls file."""
        # open the workbook
        workbook = xlrd.open_workbook(xls_file)
        # split all sheets by name
        worksheets = workbook.sheet_names()
        # acquire the first one
        worksheet = workbook.sheet_by_name(worksheets[0])
        # and append it to the list of files
        self.file_contents.append(worksheet)

    def process_expression(self, expr, myfile):
        """
        This method processes an expression from file to file.
            Algorithm:
            While expression has parentheses, then
                analyze the inside-the-parentheses expression recursively
                return a new file
                modify expression
            If expression.contains(AND):
                x = process_value(expression.left_side, file)
                y = process_value(expression.right_side, x)
                return y
            elif expression.contains(OR):
                x = process_value(expression.left_side, file)
                y = process_value(expression.right_side, file)
                z = merge(x, y)
                return z
            elif expression.contains(NOT):
                x = process_value(expression, file)
                return x
        """
        # get rid of complex expressions with parenthesis
        if (expr.has("(")):
            simple_expr = self._get_inner_expression(expr)

            if "AND" in simple_expr:
                custom_file = self._AND_expr(simple_expr, myfile)
            if "OR" in simple_expr:
                custom_file = self._OR_expr(simple_expr, myfile) #TODO write this.
            if "NOT" in simple_expr:
                custom_file = self._NOT_expr(simple_expr, myfile) #TODO write this.

    def _get_inner_expression(self, expr):
        """ This method extracts the innermost expression of
        parentheses """
        left_inner = expr.split("(")[-1]
        right_inner = expr.split(")")[0]
        simple_expr = left_inner + right_inner
        return simple_expr

    def _AND_expr(self, expr, f):
        left = simple_expr.split("AND")[0]
        right = simple_expr.split("AND")[-1]

        custom_file = self.process_file(left, f)
        custom_file = self.process_file(right, custom_file)
        return custom_file

    def process_file(self, value):
        """ This method processes a file and extracts only the data
        the user is interested for.
        """
        # open a new workbook to write the result there
        workbook = xlwt.Workbook()
        write_sheet = workbook.add_sheet('Results', cell_overwrite_ok=True)
        # index which increases depending on the number of writes
        # (in order to avoid double writes)
        max_index = 0
        # search a sheet by row (usually names are there).
        for sheet in self.file_contents:
            for row_index in range(sheet.nrows):
                row = sheet.row(row_index)
                # If the value is found, then
                # save all row
                for col_index, cell in enumerate(row):
                    if value not in str(cell.value):
                        continue
                    else:
                        write_index = self._update_write_index(max_index, row_index)
                        self._write_row(row, write_index, write_sheet)
        workbook.save(self.dest_file)

    def _write_row(self, row, row_index, write_sheet):
        """ This method writes a row of data to the new excel file"""
        for col_index, cell in enumerate(row):
            write_sheet.write(row_index, col_index, str(cell.value))

    def _update_write_index(self, max_index, row_index):
        max_index += row_index
        return max_index
