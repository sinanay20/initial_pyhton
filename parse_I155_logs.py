#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Name: parse_I155_logs.py
# Author: Samuel Bramm
# Date: 2018-01-04
# Description: Parse I155 Logs and create error files
# Version: v1.0
# Changelog
# TODO  v1.1 integrate single file analysis
# v1.0 Integrate Script into iCON scripts

from Tkinter import *
from ttk import *
from tkFileDialog import askdirectory
import re
import os
import tkMessageBox
import codecs
import time
import csv

root_window = Tk()
folderpath = ""
error_list = []
registration_list = []


def write_files():
    global error_list
    global registration_list
    global folderpath
    error_filename = folderpath + "/error_" + time.strftime("%Y_%m_%d") + ".csv"
    registration_filename = folderpath + "/registrations_" + time.strftime("%Y_%m_%d") + ".csv"
    if os.path.isfile(error_filename):
        os.remove(error_filename)
    if os.path.isfile(registration_filename):
        os.remove(registration_filename)
    with open(error_filename, 'ab') as error_file:
        error_writer = csv.writer(error_file, delimiter=";")
        error_writer.writerow(['Date', 'CountryCode', 'ContractNumber', 'ErrorCode', 'ErrorMessage'])
        for error in error_list:
            error_writer.writerow(error)
    error_file.close()
    with open(registration_filename, 'ab') as registration_file:
        registration_writer = csv.writer(registration_file, delimiter=";")
        registration_writer.writerow(['Date', 'CountryCode', 'ContractNumber', 'RegistrationMark', 'oldFIN', 'newFIN'])
        for registration in registration_list:
            registration_writer.writerow(registration)
    registration_file.close()
    return None


def parse_file(filename):
    global error_list
    global registration_list
    global folderpath
    # Search pattern definitions
    mandant_pattern = re.compile(r"ICON-Mandant    (\d{5})[A-Z]{2}")
    contract_nr_pattern = re.compile(r"Contract-Number :([A-Z0-9]{8}/\d{6}).*")
    error_pattern = re.compile(r".*<ERROR>.*-(\d{3,4})(.*)")
    reject_pattern = re.compile(r".*<REJECT>.*-(\d{3,4})(.*)")
    registration_pattern = re.compile(
        r"Registration mark ([A-Z0-9]*) is moved from FIN: ([A-Z0-9]*) to FIN: ([A-Z0-9]*).")
    old_new_registration_pattern = re.compile(r">>  FIN - OLD :\s+([A-Z0-9]*)\s+FIN - NEW :\s+([A-Z0-9]*)")
    country_pattern = re.compile(r"[Country|PAYS] (\d{5})")
    country = ''
    contract_nr = ''
    year = filename.split(".")[3][:4]
    month = filename.split(".")[3][4:6]
    day = filename.split(".")[3][6:8]
    date = day + "." + month + "." + year
    with open(folderpath + "/" + filename, 'rb') as input_file:
        for line in input_file:
            # fix encoding in file
            line = codecs.decode(line, 'cp1252')
            line = codecs.encode(line, 'utf-8')
            # search for matches in file
            if re.search(country_pattern, line):
                country = re.search(country_pattern, line).group(1)
            elif re.search(mandant_pattern, line):
                country = re.search(mandant_pattern, line).group(1)
            elif re.search(contract_nr_pattern, line):
                contract_nr = re.search(contract_nr_pattern, line).group(1)
            elif re.search(error_pattern, line):
                error_list.append([date, country, contract_nr, re.search(error_pattern, line).group(1),
                                   re.search(error_pattern, line).group(2).strip()])
            elif re.search(reject_pattern, line):
                error_list.append([date, country, contract_nr, re.search(reject_pattern, line).group(1),
                                   re.search(reject_pattern, line).group(2).strip()])
            elif re.search(registration_pattern, line):
                registration_list.append([date, country, contract_nr, re.search(registration_pattern, line).group(1),
                                          re.search(registration_pattern, line).group(2),
                                          re.search(registration_pattern, line).group(3)])
            elif re.search(old_new_registration_pattern, line):
                registration_list.append(
                    [date, country, contract_nr, 'FIN changed', re.search(old_new_registration_pattern, line).group(1),
                     re.search(old_new_registration_pattern, line).group(2)])
    return None


def filter_doublicates():
    global error_list
    global registration_list
    unique_error_list = []
    unique_registration_list = []
    for error in error_list:
        if error not in unique_error_list:
            unique_error_list.append(error)
    error_list = unique_error_list
    for registration in registration_list:
        if registration not in unique_registration_list:
            unique_registration_list.append(registration)
    registration_list = unique_registration_list
    return None


def parse_folder():
    global folderpath
    if folderpath == '':
        tkMessageBox.showerror("I/O Error", "Please specify a folder.")
    index = 1
    file_count = len(os.listdir(folderpath))
    sub_window = Toplevel()
    sub_window.title("Parsing")
    Label(sub_window, text="Parsing file " + str(index) + " of " + str(file_count) + ". Please be patient").grid(row=0,
                                                                                                                 column=0,
                                                                                                                 padx=10,
                                                                                                                 pady=10)
    p = Progressbar(sub_window, orient=HORIZONTAL, length=100, maximum=file_count, mode='determinate')
    p.grid(row=1, column=0, padx=10, pady=10)
    sub_window.update()
    for file in os.listdir(folderpath):
        p.step()
        Label(sub_window, text="Parsing file " + str(index) + " of " + str(file_count) + ". Please be patient").grid(
            row=0, column=0, padx=10, pady=10)
        sub_window.update()
        parse_file(file)
        index += 1
    Label(sub_window, text="cleaning duplicates and writing output files.").grid(row=2, column=0, padx=10, pady=10)
    sub_window.update()
    filter_doublicates()
    write_files()
    sub_window.quit()
    tkMessageBox.showinfo("Success", "All " + str(
        file_count) + " files have been parsed.\n The output files are located in " + folderpath)
    root_window.quit()
    return None


def select_folder():
    global folderpath
    folderpath = askdirectory()
    return None


def main():
    global root_window
    root_window.title("I155 Logparser")
    Label(root_window, text="1. Select folder containing I155 Logfiles.").grid(row=0, column=0, pady=10, padx=10)
    Button(root_window, text="Open", command=select_folder).grid(row=0, column=1, pady=10, padx=10)
    Label(root_window, text="2. Press the parse Button and wait.").grid(row=1, column=0, pady=10, padx=10)
    Button(root_window, text="parse", command=parse_folder).grid(row=1, column=1, pady=10, padx=10)
    root_window.mainloop()
    exit(0)


if __name__ == '__main__':
    main()
