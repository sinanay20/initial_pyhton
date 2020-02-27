#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Name: TabToComma
# Author: Benjamin Schmelcher
# Date: 20181116
# Description: Changes separator of csv file or splits large file into small ones
# Version: 1.0.0

import csv
import codecs
import os
import sys
import tkMessageBox
from Tkinter import *
from tkFileDialog import askopenfilename
from ttk import *

list = []

file_path = ""
partnumber = 0
filesize = 500000000
output_path = ""
csv.field_size_limit(sys.maxsize)
root = Tk()
globaltype = ""
selection_set = IntVar()
done = IntVar()
path_selected = IntVar()
new_operator = StringVar()
operator_set = IntVar()


def split_csv():
    global partnumber
    global output_path
    print file_path
    print output_path
    with open(file_path, 'rb') as input_file:
        partnumber = 0
        output = open(output_path+"_"+str(partnumber), 'wb')
        for line in input_file:
            output.write(line)
            output.write("\n")
            if getSize(output) >= filesize:
                output.close()
                partnumber += 1
                #output_path = output_path+str(partnumber)
                output = open(output_path+"_"+str(partnumber), 'wb')
    output.close()
    done.set(1)


def getSize(fileobject):
    fileobject.seek(0, 2)  # move the cursor to the end of the file
    size = fileobject.tell()
    return size


def tab_to_comma_output():
    with open(file_path, 'rb') as input_file:
        output = open(output_path, 'wb')
        for line in input_file:
            output.write(new_operator.join(line.split()))
            output.write("\n")
            output.flush()
    output.close()
    done.set(1)


def set_type(localtype):
    global globaltype
    globaltype = localtype
    if (globaltype == "split"):
        globaltype = "split"
    if (globaltype == "change_separator"):
        globaltype = "change_separator"
    selection_set.set(1)


def select_path():
    global output_path
    global file_path
    file_path = askopenfilename(initialdir=os.getcwd(),
                                title="Choose the *.csv file")  # , filetypes=[("CSV", "*.csv")])
    output_path = os.path.dirname(file_path) + "/output"
    if os.path.isdir(output_path):
        pass
    else:
        os.mkdir(output_path)
    output_path = output_path+"/" + os.path.basename(file_path)

    path_selected.set(1)


def main():
    root.title("Split/Change Separator of *.csv file")
    path_label = Label(root, text="Select the *.csv file")
    path_label.grid(row=1, column=1, padx=10, pady=10)
    path_button = Button(root, text="select", command=lambda: select_path())
    path_button.grid(row=1, column=2, padx=10, pady=10)
    root.wait_variable(path_selected)
    path_button.config(state=DISABLED)

    choose_label = Label(root, text="Choose between splitting and changing the separator of the *.csv file:")
    choose_label.grid(row=2, column=1, padx=10, pady=10)
    split_button = Button(root, text="Split", command=lambda: set_type("split"))
    split_button.grid(row=3, column=1, padx=10, pady=10)
    change_button = Button(root, text="Change Separator", command=lambda: set_type("change_separator"))
    change_button.grid(row=3, column=2, padx=10, pady=10)

    root.wait_variable(selection_set)
    split_button.config(state=DISABLED)
    change_button.config(state=DISABLED)

    if (globaltype == "change_separator"):
        subwindow = Toplevel()
        subwindow.title = "Change Separator of *.csv file"
        separator_label = Label(subwindow, text="Enter the new separator: ")
        separator_label.grid(row=1, column=1, padx=10, pady=10)
        separator_entry = Entry(subwindow, textvariable=new_operator)
        separator_entry.grid(row=1, column=2, padx=10, pady=10)
        separator_ok_button = Button(subwindow, text="Ok", command=lambda: operator_set.set(1))
        separator_ok_button.grid(row=2, column=2, padx=10, pady=10)
        subwindow.wait_variable(operator_set)
        separator_ok_button.config(state=DISABLED)
        parse_label = Label(subwindow, text="Creating output file...")
        parse_label.grid(row=2, column=1, padx=10, pady=10)

        tab_to_comma_output()
        subwindow.wait_variable(done)

    if (globaltype == "split"):
        parse_label = Label(root, text="Creating splitted files...")
        parse_label.grid(row=4, column=1, padx=10, pady=10)
        split_csv()
        root.wait_variable(done)

    if tkMessageBox.askyesno("Success",
                             "The Output file was created successfully. The file(s) is located under {}. Do you want to quit?".format(
                                 os.path.dirname(output_path))):
        exit(0)

    root.mainloop()
    exit(0)


if __name__ == '__main__':
    main()
