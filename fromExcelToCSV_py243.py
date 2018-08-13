#! /usr/bin/env python
# -*- coding: utf-8 -*-

""" Tool to Convert Excel file to CSV
# exceltocsv_py24.py
A terminal script written in Python that converts Excel files to CSV.
Developed using Visual Studio Code with Python extension.
Based from https://dzone.com/articles/using-python-to-extract-excel-spreadsheet-into-csv

# This script targets python 2.4.3

## Source file requirements
1. Will only accept *.xls Excel files. Not *.xlsx.
2. Source must be placed on the same directory as the python script
3. To modify default filename, update src_filename variable
4. First sheet name must be a month name (e.g. January)
5. Year must be placed at 'D3' cell (e.g. 2018)
6. Type of flight must be placed at 'C' cell on any row (e.g. International)
7. Data rows are read from 'A' to 'F' cells where 'A' cell values must be numeric

## How to run:
$ python exceltocsv.py

## Output: csv and log files
YYYYMMsrc_filename.csv
(e.g. 201801FREQUENT_TRAVELER_MONTHLY.csv)

YYYYMMDD_HHMMHH_exceltocsv.py.log
(e.g. 20180805_070007_exceltocsv.py.log)
"""

import csv
import glob
import os
import sys
import logging
from datetime import timedelta, datetime  # to calculate runtime
import time

# 3rd party module
# from xlrd import open_workbook              # to read Excel worksheets
#                                             # https://pypi.org/project/xlrd
#                                             # http://xlrd.readthedocs.io

package_path = 'packages\\xlrd-0.7.9'
sys.path.append(os.path.join(os.getcwd(), package_path))
import xlrd


# format month name
def getnumericmonth(monthname):
    month = monthname.upper()
    if month == 'JAN' or month == 'JANUARY':
        return '01'
    elif month == 'FEB' or month == 'FEBRUARY':
        return '02'
    elif month == 'MAR' or month == 'MARCH':
        return '03'
    elif month == 'APR' or month == 'APRIL':
        return '04'
    elif month == 'MAY':
        return '05'
    elif month == 'JUN' or month == 'JUNE':
        return '06'
    elif month == 'JULY' or month == 'JULY':
        return '07'
    elif month == 'AUG' or month == 'AUGUST':
        return '08'
    elif month == 'SEP' or month == 'SEPTEMBER':
        return '09'
    elif month == 'OCT' or month == 'OCTOBER':
        return '10'
    elif month == 'NOV' or month == 'NOVEMBER':
        return '11'
    elif month == 'DEC' or month == 'DECEMBER':
        return '12'
    else:
        return '00'


# main definition
def main():
    src_filename = 'food_menu.xls'
    f = glob.glob(os.path.join(src_filename))

    if not f:
        msg = 'Nothing to process. Cannot find %s file.' % src_filename
        print('%s: %s\n' % (logging.info.__name__.upper(), msg))
        logging.info(msg)
        sys.exit()
    else:
        wb = xlrd.open_workbook(src_filename)
        sheet = wb.sheet_by_index(0)
        reportyear = int(sheet.cell_value(2, 3))
        reportmonth = getnumericmonth(sheet.name)
        if reportmonth == '00':
            msg = 'Undefined month as \'%s\'. Please check source file.' % sheet.name
            print('%s: %s\n' % (logging.warning.__name__.upper(), msg))
            logging.warning(msg)
            sys.exit()
        else:
            reportmonthname = sheet.name

        msg = 'Reading %s for %s %s' % (src_filename, reportmonthname.title(), reportyear)
        print('%s: %s' % (logging.info.__name__.upper(), msg))
        logging.info(msg)

        out_filename = '%s%s%s.csv' % (reportyear, reportmonth, os.path.splitext(src_filename)[0])

        f = open(out_filename, "wb")
        try:
            writer = csv.writer(f, delimiter="|")

            typeflight = ''
            row = []
            rowcount = 0
            for row_idx in range(sheet.nrows):
                # get flight type
                if not isinstance(sheet.cell_value(row_idx, 2), str):
                    if sheet.cell_value(row_idx, 2).upper() == 'DOMESTIC':
                        typeflight = 'DOMESTIC'
                    elif sheet.cell_value(row_idx, 2).upper() == 'INTERNATIONAL':
                        typeflight = 'INTERNATIONAL'
                # fetching data
                if isinstance(sheet.cell_value(row_idx, 0), (int, float)):
                    for col_idx in range(1, 6):
                        row.insert(col_idx, sheet.cell_value(row_idx, col_idx))
                    row.insert(0, sheet.name[0:3].upper())
                    row.insert(1, reportyear)
                    row.insert(2, typeflight)
                    writer.writerow(row)

                    rowcount += 1

                    msg = 'Fetching \'%s\' data from row %s' % (typeflight.title(), row_idx)
                    print('%s: %s' % (logging.info.__name__.upper(), msg))
                    logging.info(msg)
                row = []
        finally:
            f.close()

        # summary
        if rowcount > 0:
            msg = 'Done copying %s row(s) to %s' % (rowcount, out_filename)
        else:
            msg = 'No rows to fetch.'

        elapsedtime = timedelta(seconds=round(time.time() - starttime_main))
        print('\nSUMMARY: %s | Elapsed time %s\n' % (msg, elapsedtime))
        logging.info('%s | Elapsed time %s' % (msg, elapsedtime))


# main
if __name__ == "__main__":
    # set runtime
    starttime_main = time.time()

    # set global logging
    logfilename = '%s %s.log' % (time.strftime('%Y%m%d_%H%M%S'), os.path.basename(__file__))
    logging.basicConfig(filename=os.path.join(logfilename),
                        level=logging.DEBUG,
                        format="[%(levelname)s] : %(asctime)s : %(message)s")

    # show execution start
    msg = 'Starting execution of'
    print('\n%s: %s %s' % (logging.info.__name__.upper(), msg, os.path.basename(__file__)))
    logging.info('%s %s' % (msg, os.path.basename(__file__)))

    main()
