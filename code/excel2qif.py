#!/usr/bin/env python3.6



import xlrd
import os.path
import datetime
import argparse
from argparse import RawTextHelpFormatter


parser = argparse.ArgumentParser(description=\
"""Converts ING\' xcel files to qif. 
    Creates a new file with same name but right extension (.qif)
    and content converted to qif format.
    Will overwrite if the .qif filename exists already.

    It works with my stuff, I hope it works with yours,
    but I do not provide any garantees of correctness.
    Use at your own risk, don't trust this blindly and keep
    an eye open. You have been warned.
"""
, formatter_class=RawTextHelpFormatter)

parser.add_argument('file_path', type=str,
                    help='path to the excel file.')

args = parser.parse_args()



full_path = args.file_path

book = xlrd.open_workbook(full_path)
file_name = os.path.basename(full_path)
path = os.path.dirname(full_path)
datemode = book.datemode

# print(file_name)

first_sheet = book.sheet_by_index(0)

#print(first_sheet.row_values(0))





fd = open(os.path.join(path, file_name.replace(".xls",".qif")), "w+")

fd.write("!Type:Bank\n")


for i in range(0, first_sheet.nrows) :
    row = first_sheet.row_values(i)
    assert(len(row) == 5)
    #[43731.0, 'CARTE 22/09/2019 MAISON FLORAN', '', -5.15, 'EUR']
    t = xlrd.xldate_as_tuple(row[0], datemode)
    d = datetime.datetime(*t)
    fd.write(d.strftime("D%d/%m/%Y\n"))

    fd.write("T%.2f\n" % row[3])
    fd.write("P%s\n" % row[1])
    fd.write("^\n")


fd.close()
print("Done.")
