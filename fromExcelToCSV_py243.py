#! /usr/bin/env python
# -*- coding: utf-8 -*-

""" Tool to Convert Excel file to CSV
# fromExcelToCSV_py243.py
A terminal script written in Python that reads Excel rows and export 
to a pipe-delimited csv file.

Developed using Visual Studio Code with Python extension.

This script targets python 2.4.3.

## Source file requirements
1. Will only accept *.xls Excel files and not *.xlsx
2. Source must be placed on the same directory as the python script
3. To modify default filename, update SRC_FILENAME variable
4. First sheet name must be a month name (e.g. January)
5. Year must be placed at 'E3' cell (e.g. 2018)
6. Type of menu must be placed at 'A' cell on any row (e.g. International)
7. Data rows are read from 'A' to 'E' cells where 'A' cell values must be numeric

## How to run:
$ python fromExcelToCSV_py243.py

## Output: csv and log files
YYYYMMSRC_FILENAME.csv
(e.g. 201804FOOD_MENU.csv)

YYYYMMDD_HHMMHH_exceltocsv.py.log
(e.g. 20180402_080012_fromExcelToCSV_py243.log)
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
def getNumMonthVal(monthName):
    month = monthName.upper()
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
    SRC_FILENAME = 'food_menu.xls'
    f = glob.glob(os.path.join(SRC_FILENAME))

    if not f:
        msg = 'Nothing to process. Cannot find %s file.' % SRC_FILENAME
        print('%s: %s\n' % (logging.info.__name__.upper(), msg))
        logging.info(msg)
        sys.exit()
    else:
        wb = xlrd.open_workbook(SRC_FILENAME)
        sheet = wb.sheet_by_index(0)
        reportYear = int(sheet.cell_value(2, 4))
        reportMonth = getNumMonthVal(sheet.name)
        if reportMonth == '00':
            msg = 'Undefined month as \'%s\'. Please check source file.' % sheet.name
            print('%s: %s\n' % (logging.warning.__name__.upper(), msg))
            logging.warning(msg)
            sys.exit()
        else:
            reportMonthName = sheet.name

        msg = 'Reading %s for %s %s' % (SRC_FILENAME, reportMonthName.title(), reportYear)
        print('%s: %s' % (logging.info.__name__.upper(), msg))
        logging.info(msg)

        OUT_FILENAME = '%s%s%s.csv' % (reportYear, reportMonth, os.path.splitext(SRC_FILENAME)[0])

        f = open(OUT_FILENAME, "wb")
        try:
            writer = csv.writer(f, delimiter="|")

            typeMenu = ''
            row = []
            rowcount = 0
            for row_id in range(sheet.nrows):
                # get type of menu
                # typeMenuVal = str(sheet.cell_value(row_id, 0))
                #if isinstance(sheet.cell_value(row_id, 0) , str):
                if sheet.cell_value(row_id, 0) == 'PLATTER':
                    typeMenu = 'PLATTER'
                elif sheet.cell_value(row_id, 0) == 'DRINKS':
                        typeMenu = 'DRINKS'
                # collect data and write to file
                if isinstance(sheet.cell_value(row_id, 0), (int, float)):
                    for col_id in range(1, 5):
                        row.insert(col_id, sheet.cell_value(row_id, col_id))
                    row.insert(0, sheet.name[0:3].upper())
                    row.insert(1, reportYear)
                    row.insert(2, typeMenu)
                    writer.writerow(row)

                    rowcount += 1

                    msg = 'Collecting data from \'%s\' at row %s' % (typeMenu.title(), row_id)
                    print('%s: %s' % (logging.info.__name__.upper(), msg))
                    logging.info(msg)
                row = []
        finally:
            f.close()

        # summary
        if rowcount > 0:
            msg = 'Done copying %s row(s) to %s' % (rowcount, OUT_FILENAME)
        else:
            msg = 'No rows to collect.'

        elapsedTime = timedelta(seconds=round(time.time() - startTime_Main))
        print('\nSUMMARY: %s | Elapsed time %s\n' % (msg, elapsedTime))
        logging.info('%s | Elapsed time %s' % (msg, elapsedTime))


# main
if __name__ == "__main__":
    # set runtime
    startTime_Main = time.time()

    # set global logging
    logFilename = '%s %s.log' % (time.strftime('%Y%m%d_%H%M%S'), os.path.basename(__file__))
    logging.basicConfig(filename=os.path.join(logFilename),
                        level=logging.DEBUG,
                        format="[%(levelname)s] : %(asctime)s : %(message)s")

    # show execution start
    msg = 'Starting execution of'
    print('\n%s: %s %s' % (logging.info.__name__.upper(), msg, os.path.basename(__file__)))
    logging.info('%s %s' % (msg, os.path.basename(__file__)))

    main()
