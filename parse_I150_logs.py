#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Name: parse_I150_logs.py
# Author: Samuel Bramm
# Date: 2018-01-04
# Description: Parse I150 logfiles and generate error files
# Version: v1.0
# Changelog
# v1.0 Integrate script into iCON scripts


from Tkinter import *
from ttk import *
from tkFileDialog import askopenfilenames
import tkMessageBox
import re
import os
import time
import csv
root_window = Tk()
sub_window = Toplevel()
sub_window.iconify()
file_paths = []
error_list = []

def write_file():
    global file_paths
    global error_list
    error_filename = os.path.dirname(file_paths[0]) + "/error_" + time.strftime("%Y_%m_%d") + ".csv"
    if os.path.isfile(error_filename):
        os.remove(error_filename)
    with open(error_filename, 'ab') as error_file:
        error_writer = csv.writer(error_file,delimiter=";")
        error_writer.writerow(['claim number', 'error code', 'error message'])
        for error in error_list:
            error_writer.writerow(error)
    error_file.close()
    return None

def analyze_file(filename):
    global error_list
    global sub_window
    claim_pattern = re.compile(r"ICON-CLAIM:  ([0-9]+)")
    error_pattern = re.compile(r"\s+\d{2}\s+\d*\s+-(\d{4})\s+(.*)")
    claim_number = ""
    with open(filename, 'r') as log_file:
        for line in log_file:
            sub_window.update()
            if re.search(claim_pattern, line):
                claim_number = re.search(claim_pattern, line).group(1)
            elif re.search(error_pattern, line):
                error_code = re.search(error_pattern, line).group(1)
                error_msg = re.search(error_pattern, line).group(2)
                error_list.append([claim_number,error_code,error_msg.strip()])
    return None

def parse_files():
    global file_paths
    #sub_window = Toplevel()
    global sub_window
    sub_window.deiconify()
    file_count = len(file_paths)
    #analyze each file
    p = Progressbar(sub_window, orient=HORIZONTAL, length=100, mode='indeterminate')
    p.grid(row=0,column=1,padx=10,pady=10)
    p.start()
    for i,filename in enumerate(file_paths):
        Label(sub_window, text="Analyzing file " + str(i + 1) + " of " + str(file_count) + ". Please be patient!").grid(row=0,column=0,padx=10,pady=10)
        sub_window.update()
        analyze_file(filename)
    Label(sub_window,text="All files analyzed. Writing output file.").grid(row=2,column=0,pady=10,padx=10)
    sub_window.update()
    p.stop()
    #write errors into output file
    write_file()
    tkMessageBox.showinfo("Success", "All " + str(file_count) + " file have been analyzed.\n The output file is located here in the same directory" )
    root_window.quit()
    return None

def select_file():
    global file_paths
    file_paths = askopenfilenames()
    return None

def main():
    global root_window
    root_window.title("I150 Logparser")
    Label(root_window,text="1. Select I150 Logfile").grid(row=0,column=0,pady=10,padx=10)
    Button(root_window,text="Open",command=select_file).grid(row=0,column=1,pady=10,padx=10)
    Label(root_window,text="2. Press the parse Button and wait.").grid(row=1,column=0,pady=10,padx=10)
    Button(root_window,text="parse",command=parse_files).grid(row=1,column=1,pady=10,padx=10)
    root_window.mainloop()
    exit(0)


if __name__ == '__main__':
    main()