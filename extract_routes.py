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

class Extraction:
    """ Extracts data from several files to a new one.
    """

    def __init__(self, src_folder, dest_file):
        self.file_contents = []
        self.src_folder = src_folder
        self.dest_file = dest_file
        # read the file
        self.read_folder(src_folder)
        # process data
        self.process_file()
        # and write data to a new file
        self.write_output(dest_file)

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

    def process_file(self, key="ACC ON"):
        """ This method processes a file and extracts only the data
        the user is interested for.
        """
        # open a new workbook to write the result there
        workbook = xlwt.Workbook()
        write_sheet = workbook.add_sheet('Results')
        # search a sheet by row (usually names are there).
        for sheet in self.file_contents:
            for row_index in range(sheet.nrows):
                row = sheet.row(row_index)
                # If the value is found, then
                # save all row
                for col_index, cell in enumerate(row):
                    if key not in str(cell.value):
                        continue
                    else:
                        self._write_row(row, row_index, write_sheet)
        workbook.save(self.dest_file)

    def _write_row(self, row, row_index, write_sheet):
        """ This method writes a row of data to the new excel file"""
        for col_index, cell in enumerate(row):
            write_sheet.write(row_index, col_index, str(cell.value))

    def extract_routes(self):
        pass

    def extract_parts_of_interest(self):
        pass

    def write_output(self, write_filename=""):
        """ This method writes the contents to a new file.

        :param str write_filename: the filename of the file to
        be written.
        """
        pass

def main():
    src = str(sys.argv[1])
    dest = str(sys.argv[2])
    Extraction(src, dest)

if __name__ == "__main__":
    main()
