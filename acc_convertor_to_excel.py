#-------------------------------------------------------------------------------
# Name:        PSS/E ACC to Excel
# Version:     0.4
# Purpose:
#
# Author:      oskars-g
#
# Created:     20.06.2023
# Copyright:   (c) oskars-g 2023
# Licence:     Free to modify and use
#-------------------------------------------------------------------------------

from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename
import os, sys
import psse3503
import psspy
import pssexcel

PSSE_PATH = r'C:\Program Files\PTI\PSSE35\35.3\PSSBIN'
sys.path.append(PSSE_PATH)
os.environ['PATH'] += ';' + PSSE_PATH

psspy.psseinit(7000)

def select_file(file_type='.acc', dialogue_title="Select file"):
    """ Single file selection popup"""
    #Tk().withdraw()
    root = Tk()
    root.withdraw()
    filename = askopenfilename(filetypes=[('{} file'.format(file_type), '*{}'.format(file_type))],
                               title=dialogue_title)
    return filename

if __name__ == "__main__":
    data_file = select_file()
    if data_file:
        acc_data = data_file
        pssexcel.accc(acc_data, ['s', 'e', 'v', 'b', 'l', 'g'], ratecon='a',  # export a summary
                    xlsfile= 'Contingency_Report.xls', overwritesheet=True, show=False)  # export to this file

        messagebox.showinfo("Process Completed", "Contingency report generated successfully.")
    else:
        messagebox.showwarning("No File Selected", "No file was selected. The process will now end.")
        sys.exit(0)


