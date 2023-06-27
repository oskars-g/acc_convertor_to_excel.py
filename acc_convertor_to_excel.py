#-------------------------------------------------------------------------------
# Name:        PSS/E ACC to Excel
# Version:     0.3
# Purpose:
#
# Author:      oskars-g
#
# Created:     20.06.2023
# Copyright:   (c) oskars-g 2023
# Licence:     Free to modify and use
#-------------------------------------------------------------------------------

from tkinter import *
from tkinter.filedialog import askopenfilename
import os, sys
PSSE_PATH = r'C:\Program Files\PTI\PSSE35\35.3\PSSBIN'
sys.path.append(PSSE_PATH)
os.environ['PATH'] += ';' + PSSE_PATH
import psse3503
import psspy
import pssexcel
psspy.psseinit(7000)

def select_file(file_type='.acc', dialogue_title="Select file"):
    """ Single file selection popup
    return: list"""
    Tk().withdraw()
    filename = askopenfilename(filetypes=[('{} file'.format(file_type), '*{}'.format(
        file_type))])
    return [filename]

if __name__ == "__main__":
    data_file = select_file()
    acc_data = data_file[0]
    pssexcel.accc(acc_data, ['s', 'e', 'v', 'b', 'l', 'g'], ratecon='a',  # export a summary
                    xlsfile= 'Contingency_Report.xls', overwritesheet=True, show=False)  # export to this file

