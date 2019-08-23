#! /usr/bin/env python
# -*- coding: utf-8 -*-

""" Tool that fetch rows from Excel and export to CSV file
# fromExcelToCSV.py
A terminal script written in Python that reads Excel rows and export
to a pipe-delimited csv file.

Developed using Visual Studio Code with Python extension.

## Source file requirements
1. Source must be placed on the same directory as the python script
2. To modify default filename, update SRC_FILENAME variable
3. First sheet name must be a month name (e.g. January)
4. Year must be placed at 'E3' cell (e.g. 2018)
5. Type of menu must be placed at 'A' cell on any row (e.g. International)
6. Data rows are read from 'A' to 'E' cells where 'A' cell values must be numeric

## How to run
$ python fromExcelToCSV.py

## Output: csv and log files
YYYYMMSRC_FILENAME.csv
(e.g. 201804FOOD_MENU.csv)

YYYYMMDD_HHMMHH_fromExcelToCSV.py.log
(e.g. 20180402_080012_fromExcelToCSV.py.log)
"""

import csv
import glob
import logging
import os
import sys
from datetime import timedelta, datetime    # to calculate runtime
import time

# 3rd party module
package_path = 'xlrd-1.1.0'
sys.path.append(os.path.join(os.getcwd(), 'packages', package_path))
import xlrd

# from xlrd import open_workbook              # to read Excel worksheets
#                                             # https://pypi.org/project/xlrd
#                                             # http://xlrd.readthedocs.io

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
    SRC_FILENAME = 'food_menu.xlsx'
    f = glob.glob(os.path.join(SRC_FILENAME))

    if not f:
        msg = "Nothing to process. Cannot find {} file.".format(SRC_FILENAME)
        print("{}: {}\n".format(logging.info.__name__.upper(), msg))
        logging.info(msg)
        sys.exit()
    else:
        wb = xlrd.open_workbook(SRC_FILENAME)
        sheet = wb.sheet_by_index(0)
        reportYear = int(sheet.cell_value(2, 4))
        reportMonth = getNumMonthVal(sheet.name)
        if reportMonth == '00':
            msg = "Undefined month as '{}'. Please check source file.".format(sheet.name)
            print("{}: {}\n".format(logging.warning.__name__.upper(), msg))
            logging.warning(msg)
            sys.exit()
        else:
            reportMonthName = sheet.name

        msg = "Reading {} for {} {}".format(SRC_FILENAME, reportMonthName.title(), reportYear)
        print("{}: {}".format(logging.info.__name__.upper(), msg))
        logging.info(msg)

        OUT_FILENAME = "{}{}{}.csv".format(reportYear,
                                        reportMonth,
                                        os.path.splitext(SRC_FILENAME)[0].upper()
                                    )

        with open(OUT_FILENAME, "w", newline="") as f:
            writer = csv.writer(f, delimiter = "|")

            typeMenu = ''
            row = []
            rowCount = 0
            for row_id in range(sheet.nrows):
                # get type of menu
                if isinstance(sheet.cell_value(row_id, 0), str):
                    if sheet.cell_value(row_id, 0).upper() == 'PLATTER':
                        typeMenu = 'PLATTER'
                    elif sheet.cell_value(row_id, 0).upper() == 'DRINKS':
                            typeMenu = 'DRINKS'
                # collect data and write to file
                if isinstance(sheet.cell_value(row_id, 0), (int, float)):
                    for col_id in range(1, 5):
                        row.insert(col_id, sheet.cell_value(row_id, col_id))
                    row.insert(0, sheet.name[0:3].upper())
                    row.insert(1, reportYear)
                    row.insert(2, typeMenu)
                    writer.writerow(row)

                    rowCount += 1

                    msg = "Collecting data from '{}' at row {}".format(typeMenu.title(), row_id)
                    print("{}: {}".format(logging.info.__name__.upper(), msg))
                    logging.info(msg)
                row = []

        # summary
        if rowCount > 0:
            msg = "Done copying {} row(s) to {}".format(rowCount, OUT_FILENAME)
        else:
            msg = "No rows to collect."

        elapsedTime = timedelta(seconds = round(time.time() - startTime_Main))
        print("\nSUMMARY: {} | Elapsed time {}\n".format(msg, elapsedTime))
        logging.info("{} | Elapsed time {}".format(msg, elapsedTime))

# main
if __name__ == "__main__":
    # set runtime
    startTime_Main = time.time()

    # set global logging
    logFilename = "{}_{}.log".format(time.strftime('%Y%m%d_%H%M%S'), os.path.basename(__file__))
    logging.basicConfig(filename = os.path.join(logFilename),
        level = logging.DEBUG,
        format = "[%(levelname)s] : %(asctime)s : %(message)s")

    # show execution start
    msg = 'Starting execution of'
    print("\n{}: {} {}".format(logging.info.__name__.upper(), msg, os.path.basename(__file__)))
    logging.info("{} {}".format(msg, os.path.basename(__file__)))

    main()
