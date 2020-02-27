#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Name: migrated_data_test.py
# Author: Samuel Bramm
# Co-Author: Benjamin Schmelcher
# Date: 2018-06-13
# Description: Finds contracts for given Test variants
# Version: v1.2
# Changelog
# v1.2 Changed db connection to input instead of yml
# v1.1 Added random sample option
# v1.0 Initial Release


import xlwings as xw
import yaml
from Tkinter import *
from ttk import *
from tkFileDialog import askopenfilename
import tkMessageBox
from openpyxl import *
import time
import datetime
import random
from patterns.db_connection import db_connection

root = Tk()
config_file_path = "./config/migrated_data_tests.yml"
input_file_path = ""
tenants = []
invoice_mapping = {}
pricemodel_mapping = {}
customer_mapping = {}
file_set = IntVar()
get_done = IntVar()
connection_established = IntVar()
found_contracts = []

connection = db_connection()


def write_data():
    # write found contracts to excel file
    global input_file_path
    global found_contracts
    out_file_path = os.path.dirname(input_file_path) + "/Functional_Migration_Tests_" + time.strftime(
        "%Y_%m_%d") + ".xlsx"

    try:
        book = xw.Book(input_file_path)
        sheet = book.sheets[3]

        for idx, contracts in enumerate(found_contracts):
            # print contracts
            if contracts == "":
                sheet.cells(idx + 2, 10).value = ""
            else:
                sheet.cells(idx + 2, 10).value = ", ".join(contracts)
        book.save(out_file_path)
    except IOError:
        tkMessageBox.showerror("I/O Error", "Error opening file %s" % input_file_path)

    return out_file_path


def exec_sql_stmt(stmt, num, random_pull, dbg=0):
    # execute database query and return num results

    global connection

    if dbg == 1:
        print("Executing:", stmt)
        print("Searching for %d contracts" % num)

    tmpresults = connection.execute_query(stmt)

    results = []
    for entry in tmpresults:
        results.append(entry[0])

    if random_pull == "1" and len(results) >= num:
        return random.sample(results, num)
    else:
        return results[:num]


def query_data(count, tenant, test_variants, button, date_check, number_check, random_pull):
    # create sql statements and execute them.
    global get_done
    global found_contracts
    global root
    button.config(state=DISABLED)

    for idx, test in enumerate(test_variants):
        Label(root, text="Fetching contracts for Testcase {} of {}".format(idx + 1, len(test_variants))).grid(
            row=8, column=0, padx=10, pady=10)
        root.update()
        num = count
        tmp_results = []
        if number_check.get() == "1":
            query_end = " and vc.NUMBER like '00%';"
        else:
            query_end = ";"
        if date_check.get() == "1":
            last_month = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime("%Y%m%d")
            date_end = " and vcss.PLANNEDCONTRACTEND < \'" + last_month + "\'" + query_end
            date_sql = "select vc.NUMBER, vcvs.CONTRACTSTATE, vcss.PRICEMODEL, vcss.CREDITCARDPAYMENTMETHOD, vcss.FK_PRODUCT, vcss.PAYMENTINTERVAL, pp.CUSTOMERINVOICEVARIANT, p.PARTNERTYPE from icon.CONTRACT_VEHICLECONTRACTSTABLESTATE vcss inner join icon.CONTRACT_VEHICLECONTRACT vc on vc.ACTIVESTABLESTATE_OBJECTID = vcss.OBJECTID inner join icon.CONTRACT_VEHICLECONTRACTVOLATILESTATE vcvs on vc.ACTIVEVOLATILESTATE_OBJECTID = vcvs.OBJECTID inner join icon.PRODUCT_PRODUCT pp on pp.CODE = vcss.FK_PRODUCT and vc.MASTERDATARELEASEVERSION = pp.RELEASEVERSION inner join icon.CONTRACT_CUSTOMERCONTRACTSTATE_ATTR_VEHICLECONTRACT ccsavc ON vc.NUMBER = ccsavc.VEHICLECONTRACT_NUMBER inner join icon.CONTRACT_CUSTOMERCONTRACTSTATE cccs ON ccsavc.CUSTOMERCONTRACTSTATE_OBJECTID = cccs.OBJECTID inner join icon.CONTRACT_CUSTOMERCONTRACT ccc ON cccs.CUSTOMERCONTRACT_NUMBER = ccc.NUMBER inner join icon.PARTNER_PARTNER p ON ccc.CONTRACTINGCUSTOMER_NUMBER = p.NUMBER where vcvs.CONTRACTSTATE = 'contractActive' and vc.TENANTID=\'" + tenant + "\' and vcss.TENANTID=\'" + tenant + "\' and vcvs.TENANTID=\'" + tenant + "\' and pp.TENANTID=\'" + tenant + "\' and ccsavc.TENANTID=\'" + tenant + "\' and cccs.TENANTID=\'" + tenant + "\' and ccc.TENANTID=\'" + tenant + "\' and p.TENANTID=\'" + tenant + "\' and vcss.FK_PRODUCT=\'" + \
                       test[0] + "\' and vcss.PRICEMODEL" + test[1] + " and p.PARTNERTYPE" + test[
                           3] + " and vcss.PAYMENTINTERVAL" + test[2][
                           "payment_interval"] + " and pp.CUSTOMERINVOICEVARIANT" + test[2][
                           "customer_invoice_variant"] + " and COALESCE(vcss.CREDITCARDPAYMENTMETHOD,'')" + test[2][
                           "credit_card_payment_method"] + date_end
            date_results = exec_sql_stmt(date_sql, 1, random_pull)
            if len(date_results) == 1:
                num -= 1
                tmp_results.append(date_results[0])
        test_sql = "select vc.NUMBER, vcvs.CONTRACTSTATE, vcss.PRICEMODEL, vcss.CREDITCARDPAYMENTMETHOD, vcss.FK_PRODUCT, vcss.PAYMENTINTERVAL, pp.CUSTOMERINVOICEVARIANT, p.PARTNERTYPE from icon.CONTRACT_VEHICLECONTRACTSTABLESTATE vcss inner join icon.CONTRACT_VEHICLECONTRACT vc on vc.ACTIVESTABLESTATE_OBJECTID = vcss.OBJECTID inner join icon.CONTRACT_VEHICLECONTRACTVOLATILESTATE vcvs on vc.ACTIVEVOLATILESTATE_OBJECTID = vcvs.OBJECTID inner join icon.PRODUCT_PRODUCT pp on pp.CODE = vcss.FK_PRODUCT and vc.MASTERDATARELEASEVERSION = pp.RELEASEVERSION inner join icon.CONTRACT_CUSTOMERCONTRACTSTATE_ATTR_VEHICLECONTRACT ccsavc ON vc.NUMBER = ccsavc.VEHICLECONTRACT_NUMBER inner join icon.CONTRACT_CUSTOMERCONTRACTSTATE cccs ON ccsavc.CUSTOMERCONTRACTSTATE_OBJECTID = cccs.OBJECTID inner join icon.CONTRACT_CUSTOMERCONTRACT ccc ON cccs.CUSTOMERCONTRACT_NUMBER = ccc.NUMBER inner join icon.PARTNER_PARTNER p ON ccc.CONTRACTINGCUSTOMER_NUMBER = p.NUMBER where vcvs.CONTRACTSTATE = 'contractActive' and vc.TENANTID=\'" + tenant + "\' and vcss.TENANTID=\'" + tenant + "\' and vcvs.TENANTID=\'" + tenant + "\' and pp.TENANTID=\'" + tenant + "\' and ccsavc.TENANTID=\'" + tenant + "\' and cccs.TENANTID=\'" + tenant + "\' and ccc.TENANTID=\'" + tenant + "\' and p.TENANTID=\'" + tenant + "\' and vcss.FK_PRODUCT=\'" + \
                   test[0] + "\' and vcss.PRICEMODEL" + test[1] + " and p.PARTNERTYPE" + test[
                       3] + " and vcss.PAYMENTINTERVAL" + test[2][
                       "payment_interval"] + " and pp.CUSTOMERINVOICEVARIANT" + test[2][
                       "customer_invoice_variant"] + " and COALESCE(vcss.CREDITCARDPAYMENTMETHOD,'')" + test[2][
                       "credit_card_payment_method"] + query_end
        test_results = exec_sql_stmt(test_sql, num, random_pull)
        for r in test_results:
            tmp_results.append(r)
        if len(tmp_results) == 0:
            found_contracts.append("")
        else:
            found_contracts.append(tmp_results)
    get_done.set(1)
    return None


def sanitize_variants(test_variants):
    # map excel values to db values for test variants
    global pricemodel_mapping
    global invoice_mapping
    global customer_mapping
    for idx, test in enumerate(test_variants):
        test_variants[idx][1] = pricemodel_mapping[test[1]]
        test_variants[idx][2] = invoice_mapping[test[2]]
        test_variants[idx][3] = customer_mapping[test[3]]
    return test_variants


def read_excel_file():
    # read excel file and return test variants
    global input_file_path
    test_variants = []
    try:
        book = load_workbook(input_file_path, data_only=True)
        sheet = book.get_sheet_by_name("Test Variants")
        for line in sheet:
            if line[2].value != "#N/A" and line[2].value is not None:
                test_variants.append([str(line[2].value), str(line[4].value), str(line[6].value), str(line[8].value)])
    except IOError:
        tkMessageBox.showerror("I/O Error",
                               "Error opening file %s. Pleas check if you have access to it." % input_file_path)
    return test_variants[1:]


def read_config():
    # read configuration data from file
    global config_file_path
    global tenants
    global invoice_mapping
    global pricemodel_mapping
    global customer_mapping
    try:
        config = yaml.safe_load(open(config_file_path))
        tenants = config["tenants"]
        pricemodel_mapping.update(config["mappings"]["pricemodel"])
        invoice_mapping.update(config["mappings"]["invoicing"])
        customer_mapping.update(config["mappings"]["customer"])
    except IOError:
        tkMessageBox.showerror("I/O Error", "Error reading configuration file from %s" % config_file_path)
    return None


def select_file():
    # ask user for filename and update global file path
    global input_file_path
    global file_set
    input_file_path = askopenfilename(initialdir=os.getcwd(), title="Choose the Excel file",
                                      filetypes=[("Excel", "*.xlsx")])
    file_set.set(1)
    return None


def connect_to_db(connect_button):
    global connection_established
    global connection

    connect_button.config(state=DISABLED)
    connection.show_db_data_window()
    connection_established.set(1)

def exit_function():
    connection.close_connection()
    os._exit(-1)


def main():
    # display windows and main functions
    global root
    global file_set
    global tenants
    global get_done
    # nr = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    nr = range(1, 15)
    root.protocol("WM_DELETE_WINDOW", exit_function)
    root.title("Migrated Data Test Generator")
    select_label = Label(root, text="1. Select Excel file containing test variants")
    select_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)
    select_button = Button(root, text="select", command=lambda: select_file())
    select_button.grid(row=0, column=2, padx=10, pady=10)
    root.wait_variable(file_set)
    select_button.config(state=DISABLED)
    test_variants = read_excel_file()
    test_variants = sanitize_variants(test_variants)

    connect_label = Label(root, text="2. Click the connect button")
    connect_label.grid(row=1, column=0, padx=10, pady=10)
    connect_button = Button(root, text="connect", command=lambda: connect_to_db(connect_button))
    connect_button.grid(row=1, column=2, padx=10, pady=10)
    root.wait_variable(connection_established)

    data_label = Label(root,
                       text="3. Specify how many contracts and for which tenant should be\nfetched and adjust the optional settings if needed.")
    data_label.grid(row=2, column=0, columnspan=3, padx=10, pady=10)
    nr_label = Label(root, text="count")
    nr_label.grid(row=3, column=0, padx=10, pady=10)
    tenant_label = Label(root, text="tenant")
    tenant_label.grid(row=3, column=1, padx=10, pady=10)
    count = IntVar()
    nr_dropdown = OptionMenu(root, count, 5, *sorted(nr))
    nr_dropdown.grid(row=4, column=0, padx=10, pady=10)
    tenant = StringVar()
    tenant_dropdown = OptionMenu(root, tenant, *sorted(tenants))
    tenant_dropdown.grid(row=4, column=1, padx=10, pady=10)
    date_check = StringVar(root)
    date_label = Label(root, text="(optional) Find contracts with enddate last month", anchor="w")
    date_label.grid(row=5, column=0, columnspan=2, padx=10, pady=10)
    date_button = Checkbutton(root, variable=date_check)
    date_button.grid(row=5, column=2, padx=10, pady=10)
    number_check = StringVar(root)
    number_label = Label(root, text="(optional) Find only contract numbers starting with 00", anchor="w")
    number_label.grid(row=6, column=0, columnspan=2, padx=10, pady=10)
    number_button = Checkbutton(root, variable=number_check)
    number_button.grid(row=6, column=2, padx=10, pady=10)

    random_pull = StringVar(root)
    random_label = Label(root, text="(optional) Take random contracts for each test variant", anchor="w")
    random_label.grid(row=7, column=0, columnspan=2, padx=10, pady=10)
    random_button = Checkbutton(root, variable=random_pull)
    random_button.grid(row=7, column=2, padx=10, pady=10)

    exec_label = Label(root, text="3. Press the get button.", anchor="w")
    exec_label.grid(row=8, column=0, columnspan=2, padx=10, pady=10)
    exec_button = Button(root, text="get",
                         command=lambda: query_data(count.get(), tenant.get(), test_variants,
                                                    exec_button, date_check, number_check, random_pull.get()))
    exec_button.grid(row=8, column=2, padx=10, pady=10)
    root.wait_variable(get_done)
    out_file = write_data()
    if tkMessageBox.askyesno("Success",
                             "The Contingent was sucessfully mapped. The Outputfile is in %s. Do you want to quit?" % out_file):
        connection.close_connection()
        exit(0)
    root.mainloop()
    connection.close_connection()
    exit(0)


if __name__ == '__main__':
    read_config()
    main()
