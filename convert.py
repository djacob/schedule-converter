#!/bin/python

import sys
import xlrd
import unicodecsv as csv

def xls_to_csv(xls_path, csv_path):
    """
    Converts the xls file at the given path and writes out a csv file
    :param xls_path the input xls file path
    :param csv_path the output csv file path
    """
    workbook = xlrd.open_workbook(xls_path)
    print('sheets {0}'.format(workbook.sheets()))
    worksheet = workbook.sheet_by_index(1)
    csvfile = open(csv_path, 'wb')
    wr = csv.writer(csvfile)

    for rownum in xrange(worksheet.nrows):
        row = []
        for cell in worksheet.row_values(rownum):
            if type(cell) == type(u''):
                row.append(cell.encode('utf-8'))
            elif type(cell) == type(0.1) and cell > 0 and cell < 1:
                time = xlrd.xldate_as_tuple(cell, datemode=1)
                row.append("{0}:{1}".format(time[3],time[4]))
            else:
                row.append(cell)
        wr.writerow(row)

    csvfile.close()

if __name__ == "__main__":
    xls_path = sys.argv[1]
    csv_path = sys.argv[2]

    if xls_path and csv_path:
        xls_to_csv(xls_path, csv_path)
    else:
        print('usage {0} xls_path csv_path'.format(sys.argv[0]))

