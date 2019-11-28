#!/usr/bin/env python3.6



import xlrd
import os.path
import datetime
import argparse
from argparse import RawTextHelpFormatter


parser = argparse.ArgumentParser(description=\
"""Converts bank transactions records to QIF files 
    For now it supports :
        ING excel files.
        SOGE csv files.
    Creates a new file with same name (ING) or new name (SOGE)
    but right extension (.qif) and content converted to qif format.
    Warning !! : Will overwrite if the .qif filename exists already.
    
    It works with my stuff, I hope it works with yours,
    but I do not provide any garantees of correctness.
    Use at your own risk, don't trust this blindly and keep
    an eye open. You have been warned.
"""
, formatter_class=RawTextHelpFormatter)

parser.add_argument('file_path', type=str,
                    help='path to the excel file.')
args = parser.parse_args()


def do_ing(full_path):
    book = xlrd.open_workbook(full_path)
    file_name = os.path.basename(full_path)
    path = os.path.dirname(full_path)
    datemode = book.datemode
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

def do_soge(full_path):
    lines = [ l for l in open(full_path,'r', encoding='iso-8859-1') ]
    path = os.path.dirname(full_path)
    file_name = os.path.basename(full_path)


    #example de filename
    #"Export_28052019_27112019.csv"

    #example de 1ere ligne
    #="0067000051772201";28/05/2019;27/11/2019;
    account_no = lines[0].split(";")[0][7:18]
    
    import datetime
    import time
    dest_fname = datetime.datetime.fromtimestamp(time.time()).strftime('%Y %m %d '+account_no + '.qif')

    fd = open(os.path.join(path, dest_fname), "w+")
    fd.write("!Type:Bank\n")
    for l in lines[2:] :
        #exemple de ligne
        #25/11/2019;000001 VIR PERM POUR: VALENTINE CURTIT REF: 1000103251016 MOTIF: LIVRET A DE VALENT;-50,00;EUR;
        items = l.split(";")
        fd.write("D%s\n" % items[0])

        fd.write("T%s\n" % items[2])
        fd.write("P%s\n" % items[1])
        fd.write("^\n")

    fd.close()


#Pour le moment, utilisons une heuristique simple ...
if args.file_path.endswith(".csv") :
    do_soge(args.file_path)
elif args.file_path.endswith(".xls") :
    do_ing(args.file_path)


print("Done.")
