#!/bin/python

import sys
import xlrd
from datetime import date as datetime

def xls_to_list(xls_path):
    """
    Converts the xls file at the given path and writes out a csv file
    :param xls_path the input xls file path
    """
    workbook = xlrd.open_workbook(xls_path)
    worksheet = workbook.sheet_by_index(1)
    sheet = []
    for rownum in xrange(worksheet.nrows):
        row = []
        for cell in worksheet.row_values(rownum):
            if type(cell) == type(u''):
                row.append(cell.encode('utf-8'))
            elif type(cell) == type(0.1) and cell > 0:
                date = xlrd.xldate_as_tuple(cell, datemode=1)
                if date[0] == 0:
                    row.append("{0}:{1}".format(date[3], date[4]))
                else:
                    row.append("{0}-{1}-{2}".format(date[1], date[2] - 1, datetime.today().year))
            else:
                row.append(cell)
        sheet.append(row)

    return sheet


def sheet_to_shifts(sheet):
    shifts = []
    dates = []
    for i, row in enumerate(sheet):
        if row[3] == 'Sunday':
            dates = sheet[i + 1]
        if row[1] == 'Meg':
            time_in = sheet[i]
            time_out = sheet[i + 1]
            prev_time = None
            for i, date in enumerate(dates):
                shift = 'first' if date else ('second' if prev_time else None)
                if time_in != '' and shift:
                    shifts.append((date if date != '' else dates[i - 1], time_in[i], time_out[i]))
                prev_time = date if date != '' else None

    return shifts


if __name__ == "__main__":
    xls_path = sys.argv[1]
    sheet = xls_to_list(xls_path=xls_path)
    shifts = sheet_to_shifts(sheet)

    for shift in shifts:
        if shift[1] != '' and shift[1] != '0:15':
            print shift
