#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Name: cofico_feedback_splitter.py
# Author: Samuel Bramm
# Date: 2018-06-15
# Description: split cofico feedback files into files containing status 10 as well as 20 & 40
# Version: v0.9
from Tkinter import *
from ttk import *
from tkFileDialog import askdirectory
import tkMessageBox
import os
from openpyxl import *


root = Tk()
select_done = IntVar()
split_done = IntVar()
cofico_file_path = ""
status_10 = []
status_20_40 = []


def write_output_files():
    # write files and return output folder path
    global status_10
    global status_20_40
    global cofico_file_path
    out_path_folder = cofico_file_path + "/output/"
    try:
        os.stat(out_path_folder)
    except:
        os.mkdir(out_path_folder)
    with open(out_path_folder + "status_10.txt", 'a') as out_file_10:
        for status in status_10:
            out_file_10.write(status)
    with open(out_path_folder + "status_20_20.txt", 'a') as out_file_20_40:
        for status in status_20_40:
            out_file_20_40.write(status)

    return out_path_folder


def split_files():
    # split files into status 10 and status 20 & 40
    global status_10
    global status_20_40
    global split_done
    global cofico_file_path
    for file in os.listdir(cofico_file_path):
        try:
            with open(cofico_file_path + "/" + file, 'r') as input_file:
                for i, line in enumerate(input_file):
                    if i > 0:
                        if line.startswith("10"):
                            status_10.append(line)
                        elif line.startswith("20"):
                            status_20_40.append(line)
                        elif line.startswith("40"):
                            status_20_40.append(line)
                        else:
                            pass
        except IOError:
            tkMessageBox.showerror("I/O Error", "Error opening file %s" % file)
    split_done.set(1)
    return None


def select_path():
    # select path containing CoFiCo files
    global cofico_file_path
    global select_done
    cofico_file_path = askdirectory(initialdir=os.getcwd(), title="Choose the folder containing the cofico files")
    select_done.set(1)
    return None


def main():
    # display main window and options
    global root
    global select_done
    global split_done
    root.title("CoFiCo Feedback Splitter")
    select_label = Label(root, text="1. Select folder containing CoFiCo Feedback Files")
    select_label.grid(row=0, column=0, padx=10, pady=10)
    select_button = Button(root, text="select", command=lambda: select_path())
    select_button.grid(row=0, column=1, padx=10, pady=10)
    root.wait_variable(select_done)
    split_label = Label(root, text="2. Press the split button.")
    split_label.grid(row=1, column=0, padx=10, pady=10)
    split_button = Button(root, text="split", command=lambda: split_files())
    split_button.grid(row=1, column=1, padx=10, pady=10)
    root.wait_variable(split_done)
    out_path = write_output_files()
    if tkMessageBox.askyesno("Sucess", "The files were splitted and the output files are in the folder %s" % out_path):
        exit(0)
    root.mainloop()
    exit(0)


if __name__ == '__main__':
    main()
