#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Name: extract_contract_numnbers
# Author: Samuel Bramm
# Co-Author: Benjamin Schmelcher
# Date: 2018-08-30
# Description: extracts contract numbers from database for failed simulated invoice run
# Version: v0.9b
from Tkinter import *
from tkFileDialog import askdirectory
from ttk import *
import tkMessageBox
import os
import time
import csv
import yaml
import re

from datetime import datetime, timedelta

from patterns.db_connection import db_connection

config_file_path = './config/extract_contract_numbers.yml'
root = Tk()
tenants = []
load_done = IntVar()
found_contract_nrs = []

connection_set = IntVar()

connection = db_connection()

def write_contract_nrs():
    global found_contract_nrs
    output_folder = askdirectory(initialdir=os.getcwd(), title="Choose output folder")
    out_file_path = output_folder+"/found_contract_numbers_{}.csv".format(time.strftime("%Y_%m_%d"))

    try:
        out_file = open(out_file_path, 'wb')
        out_writer = csv.writer(out_file, delimiter=";")
        for number in found_contract_nrs:
            out_writer.writerow([number])
        return out_file_path
    except IOError:
        tkMessageBox.showerror("I/O Error", "Error writing to file {}".format(out_file_path))
        return None


def get_data(tenant, exec_button, search_date):
    exec_button.config(state=DISABLED)
    global found_contract_nrs
    global load_done


    stmt = "SELECT ERRORMESSAGE FROM icon.COMMON_INTERFACEPROTOCOL WHERE  EXECUTIONTIME >= \'" + search_date + " 00:00:00' and INTERFACENAME='SIMULATEPERIODICINVOICES' and TENANTID=\'" + tenant + "\';"
    error_results = connection.execute_query(stmt)

    contract_nr_pattern = re.compile(r'([A-Z0-9]{8}/[A-Z0-9]{6})')

    for error_result in error_results:
        contract_nr = re.search(contract_nr_pattern, error_result[0])
        found_contract_nrs.append(contract_nr.group(1))


    load_done.set(1)
    return None


def load_config():
    global config_file_path
    global tenants
    try:
        config = yaml.load(open(config_file_path))
        tenants = config["tenants"]

    except IOError:
        tkMessageBox.showerror("I/O Error", "Error reading configuration file from %s" % config_file_path)
    return None

def show_connect_window(connection_button):
    global connection_set
    global connection

    connection_button.config(state=DISABLED)
    connection.show_db_data_window()

    connection_set.set(1)


def exit_function():
    connection.close_connection()
    os._exit(-1)


def main():
    global root
    global tenants
    global server_data
    global load_done
    root.protocol("WM_DELETE_WINDOW", exit_function)
    root.title("Extract Contract Numbers")

    connection_label = Label(root, text="1. Establish a putty connection to the database")
    connection_label.grid(row=0, column=0, padx=10, pady=10)

    connect_label = Label(root, text="2. Press the connect button")
    connect_label.grid(row=1, column=0, padx=10, pady=10)

    connect_button = Button(root, text="connect", command=lambda: show_connect_window(connect_button))
    connect_button.grid(row=1, column=1, padx=10, pady=10)

    root.wait_variable(connection_set)

    select_label = Label(root,
                         text="3. Select the desired tenant and specify a search date (YYYY-MM-DD)")
    select_label.grid(row=2, column=0, padx=10, pady=10)
    tenant_label = Label(root, text="tenant")
    tenant_label.grid(row=3, column=0, padx=10, pady=10)
    tenant = StringVar()
    tenant_dropdown = OptionMenu(root, tenant, *sorted(tenants))
    tenant_dropdown.grid(row=4, column=0, padx=10, pady=10)
    date_label = Label(root, text="date")
    date_label.grid(row=3, column=2, padx=10, pady=10)
    search_date = StringVar()
    date_entry = Entry(root, textvariable=search_date)
    date_entry.insert(END, (datetime.now() + timedelta(days=-1)).strftime("%Y-%m-%d"))
    date_entry.grid(row=4, column=2, pady=10, padx=10)
    exec_label = Label(root, text="3. Press the load button to get the data from the database.")
    exec_label.grid(row=5, column=0, padx=10, pady=10, columnspan=2)
    exec_button = Button(root, text="load",
                         command=lambda: get_data(tenant.get(), exec_button, search_date.get()))
    exec_button.grid(row=5, column=2, padx=10, pady=10)
    root.wait_variable(load_done)
    out_file_path = write_contract_nrs()
    if out_file_path is None:
        tkMessageBox.showerror("Error", "Error writing file. Abotring!")
        connection.close_connection()
        exit(1)
    else:
        if tkMessageBox.askyesno("Sucess",
                                 "The found contract_numbers have been written to the file {}. Do you want to quit?".format(
                                         out_file_path)):
            connection.close_connection()
            exit(0)
    root.mainloop()
    connection.close_connection()
    exit(0)


if __name__ == '__main__':
    load_config()
    main()
