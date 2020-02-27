#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Name: create_credit_notes.py
# Author: Samuel Bramm
# Date: 2018-01-04
# Description: Convert credit notes from csv file to xml files
# Version: v1.1
# Changelog
# v0.1 Betaversion
# v1.0 Release 2018-01-03
# v1.1 Integration into iCON Scripts

from Tkinter import *
from ttk import *
from tkFileDialog import askopenfilename
import tkMessageBox
import os
import csv
import time

root_window = Tk()
filename = ""
# optional settings
optional_parameters = [["tenantId", "53137FR"], ["causation", "operational"], ["currency", "EUR"],
                       ["customerInvoiceDefinition", "MANU"], ["documentDate", time.strftime("%Y%m01")],
                       ["financialDocumentSource", "imported"], ["status", "created"], ["externalIdPrefix", "CN_"],
                       ["sourceSystem", "xmlFiles"], ["vatAmount", "0.2"], ["creationDate", time.strftime("%Y%m01")],
                       ["taxation", "CUS_INV_PRIVATE"], ["periodFrom", time.strftime("%Y%m01")]]


def calc_amounts(price, vat):
    netAmount = price.replace(',', '.')
    netAmount = round(float(netAmount), 2)
    vatAmount = round(float(netAmount) * float(vat), 2)
    grossAmount = round(vatAmount + float(netAmount), 2)

    paymentAmount = grossAmount
    invoicePositionNetAmount = netAmount
    invoicePositionVATAmount = vatAmount

    grossAmount = str(grossAmount)
    netAmount = str(netAmount)
    vatAmount = str(vatAmount)
    paymentAmount = str(paymentAmount)
    invoicePositionVATAmount = str(invoicePositionVATAmount)
    invoicePositionNetAmount = str(invoicePositionNetAmount)
    return grossAmount, netAmount, vatAmount, paymentAmount, invoicePositionNetAmount, invoicePositionVATAmount


def update_values(new_values, window):
    global optional_parameters
    global root_window
    for i, value in enumerate(new_values):
        if value.get() != '':
            optional_parameters[i][1] = value.get()
    # read csv file
    credit, invoice = read_csv_file()
    # create xml files
    if len(credit) != 0:
        create_xml_file(credit, "creditNote")
    if len(invoice) != 0:
        create_xml_file(invoice, "invoice")
    tkMessageBox.showinfo("Success", "Sucessfully create the files in the dirctory " + os.path.dirname(filename) + ".")
    root_window.quit()
    return None


def optional_settings():
    global optional_parameters
    global filename
    optional_parameters[7][1] = optional_parameters[7][1] + os.path.basename(filename)[:5]
    advanced = Toplevel()
    advanced.title("Advanced Options")
    description = Label(advanced,
                        text="Define custom values if the default ones do not match. Leave the fields blank otherwise.")
    description.grid(row=0, column=0, columnspan=4, pady=20, padx=10)
    optional_name = Label(advanced, text="Parameter")
    optional_name.grid(row=1, column=0, pady=20, padx=10)
    optional_value = Label(advanced, text="Value")
    optional_value.grid(row=1, column=1, pady=20, padx=10)
    optional_new = Label(advanced, text="New value")
    optional_new.grid(row=1, column=2, pady=20, padx=10)
    new_values = []
    index = 2
    for i, parameter in enumerate(optional_parameters):
        Label(advanced, text=parameter[0]).grid(row=(i + 2), column=0, pady=10, padx=10)
        Label(advanced, text=parameter[1]).grid(row=(i + 2), column=1, pady=10, padx=10)
        new_values.append(StringVar())
        Entry(advanced, textvariable=new_values[i]).grid(row=(i + 2), column=2, pady=10, padx=10)
        index += 1
    Button(advanced, text="Done", command=lambda: update_values(new_values, advanced)).grid(row=index, column=2,
                                                                                            pady=20, padx=10)
    return None


def create_xml_file(list, type):
    global optional_parameters
    global filename
    os.chdir(os.path.dirname(filename))
    output_filename = type + "_08-Revenue_" + time.strftime("%Y-%m-%dT%H-%M-%S_") + optional_parameters[0][
        1] + "_iCON_tec_EDF01_TC001-operational-OK-rev70-INV.xml"
    output_writer = ""
    try:
        output_writer = open(output_filename, 'a')
    except:
        tkMessageBox.showerror("I/O Error",
                               "Can not open file" + os.path.dirname(filename) + "/" + output_filename + "for writing")
        exit(1)
    # create Header in file
    output_writer.write("<common_il:ServiceInvocationCollection xmlns:actor_pl=\"http://actor.icon.daimler.com/pl\" "
                        "xmlns:calculation_pl=\"http://calculation.icon.daimler.com/pl\" "
                        "xmlns:common_il=\"http://common.icon.daimler.com/il\" "
                        "xmlns:contract_il=\"http://contract.icon.daimler.com/il\" "
                        "xmlns:contract_pl=\"http://contract.icon.daimler.com/pl\" "
                        "xmlns:contract_sl=\"http://contract.icon.daimler.com/sl\" "
                        "xmlns:contract_ui=\"http://contract.icon.daimler.com/ui\" "
                        "xmlns:cost_pl=\"http://cost.icon.daimler.com/pl\" "
                        "xmlns:cost_sl=\"http://cost.icon.daimler.com/sl\" "
                        "xmlns:jxb=\"http://java.sun.com/xml/ns/jaxb\" xmlns:logging_pl=\"http://logging.icon.daimler.com/pl\" "
                        "xmlns:logging_sl=\"http://logging.icon.daimler.com/sl\" "
                        "xmlns:logging_ui=\"http://logging.icon.daimler.com/ui\" "
                        "xmlns:masterdata_pl=\"http://masterdata.icon.daimler.com/pl\" "
                        "xmlns:masterdata_sl=\"http://masterdata.icon.daimler.com/sl\" "
                        "xmlns:mdsd_sl=\"http://mdsd.icon.daimler.com/sl\" xmlns:partner_pl=\"http://partner.icon.daimler.com/pl\""
                        " xmlns:product_pl=\"http://product.icon.daimler.com/pl\" "
                        "xmlns:product_sl=\"http://product.icon.daimler.com/sl\" "
                        "xmlns:quotation_sl=\"http://quotation.icon.daimler.com/sl\" "
                        "xmlns:report_sl=\"http://report.icon.daimler.com/sl\" "
                        "xmlns:revenue_pl=\"http://revenue.icon.daimler.com/pl\" "
                        "xmlns:revenue_sl=\"http://revenue.icon.daimler.com/sl\" "
                        "xmlns:system_pl=\"http://system.icon.daimler.com/pl\" "
                        "xmlns:system_sl=\"http://system.icon.daimler.com/sl\" "
                        "xmlns:system_ui=\"http://system.icon.daimler.com/ui\" "
                        "xmlns:vehicle_pl=\"http://vehicle.icon.daimler.com/pl\" "
                        "xmlns:vehicle_sl=\"http://vehicle.icon.daimler.com/sl\" "
                        "xmlns:xjc=\"http://java.sun.com/xml/ns/jaxb/xjc\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n")
    output_writer.write("  <executionSettings xmlns:ns1=\"http://common.icon.daimler.com/il\" "
                        "dateTime=\" " + time.strftime(
        "%Y-%m-%dT00:00:00.000") + "\" userId=\"iCON_tec_EDF01\" tenantId=\"" + optional_parameters[0][
                            1] + "\" causation=\"" + optional_parameters[1][1] + "\" "
                                                                                 "additionalInformation1=\"1\" issueThreshold=\"error\"/>\n")
    # create entries for credit notes
    for i, entry in enumerate(list):
        grossAmount, netAmount, vatAmount, paymentAmount, invoicePositionNetAmount, invoicePositionVATAmount = calc_amounts(
            entry[2], optional_parameters[9][1])
        externalId = optional_parameters[7][1] + "_" + str(`i + 1`).zfill(5)
        output_writer.write(
            "	<invocation xmlns:ns1=\"http://common.icon.daimler.com/il\" operation=\"createCustomerFinancialDocument\">\n")
        output_writer.write(
            "		<parameter xsi:type=\"revenue_pl:RevenueType\" currency=\"" + optional_parameters[2][1] +
            "\" customerInvoiceDefinition=\"" + optional_parameters[3][1] + "\" documentDate=\"" +
            optional_parameters[4][1] +
            "\" financialDocumentDefinition=\"" + type + "\" financialDocumentSource=\"" + optional_parameters[5][1] +
            "\" grossAmount=\"" + grossAmount + "\" invoiceDate=\"" + entry[3] + "\" netAmount=\"" + netAmount +
            "\" paymentAmount=\"" + paymentAmount + "\" status=\"" + optional_parameters[6][
                1] + "\" vatAmount=\"" + vatAmount +
            "\" externalId=\"" + externalId + "\" sourceSystem=\"" + optional_parameters[8][1] + "\">\n")
        output_writer.write(
            "			<financialDocumentReceiver xsi:type=\"partner_pl:PhysicalPersonType\" number=\"" + entry[
                1] + "\"/>\n")
        output_writer.write(
            "			<position creationDate=\"" + optional_parameters[10][1] + "\" invoiceLineGroupNumber=\"1\" "
                                                                                   "invoicePositionNetAmount=\"" + invoicePositionNetAmount + "\" invoicePositionVATAmount=\"" + invoicePositionVATAmount +
            "\" periodFrom=\"" + optional_parameters[12][1] + "\" positionNumber=\"1\" quantity=\"1\" taxation=\"" +
            optional_parameters[11][1] + "\"/>\n")
        output_writer.write("		</parameter>\n")
        output_writer.write("		<parameter/>\n")
        output_writer.write("		<parameter/>\n")
        output_writer.write("		<parameter>" + entry[0] + "</parameter>\n")
        output_writer.write("	</invocation>\n")
    # finalize file
    output_writer.write("</common_il:ServiceInvocationCollection>\n")
    output_writer.close()
    return output_filename


def read_csv_file():
    credit = []
    invoice = []
    global filename
    with open(filename, 'rb') as csvfile:
        reader = csv.reader(csvfile, delimiter=";")
        for i, line in enumerate(reader):
            if i > 0:
                if "invoice" in line[4].lower():
                    invoice.append(line)
                if "credit" in line[4].lower():
                    credit.append(line)
    return credit, invoice


def check_and_create(optional_choice):
    global filename
    global root_window
    # check if file exists
    if not os.path.isfile(filename):
        tkMessageBox.showerror("I/O error", "The selected file " + filename + " does not exist or is inaccessible!")
    # check if columns match
    if len(open(filename).readline().split(";")) != 5 or "number" not in open(filename).readline().split(";")[
        0].lower():
        tkMessageBox.showerror("Data error",
                               "The contents in the file " + filename + " do not match the expected criteris.\n Please provide a csv file with the following columns:\nContract Number;Financialdocumentreceiver number;Credit amount;Invoice Date;Document type")
    if optional_choice == 1:
        optional_settings()
    else:
        # read csv file
        credit, invoice = read_csv_file()
        # create xml files
        if len(credit) != 0:
            create_xml_file(credit, "creditNote")
        if len(invoice) != 0:
            create_xml_file(invoice, "invoice")
        tkMessageBox.showinfo("Success",
                              "Sucessfully create the files in the dirctory " + os.path.dirname(filename) + ".")
        root_window.quit()
    return None


def select_file():
    global filename
    filename = askopenfilename(filetypes=[("csv file", "*.csv")])
    return None


def display_window():
    global root_window
    root_window.title("Credit Note Converter")
    main_description = Label(root_window,
                             text="The Script converts a csv-File to two seperate xml-Files for Credit Notes and Invoices.")
    main_description.grid(row=0, column=0, columnspan=2, pady=20, padx=20)
    description_label = Label(root_window,
                              text="The script expects a csv file with the following columns:\nContract Number;Financialdocumentreceiver number;Credit amount;Invoice Date;Document type")
    description_label.grid(row=1, column=0, columnspan=2, pady=20, padx=20)
    select_label = Label(root_window, text="1. Select the csv file to convert.")
    select_button = Button(root_window, text="Open", command=select_file)
    select_label.grid(row=2, column=0, pady=20, padx=20)
    select_button.grid(row=2, column=1, pady=20, padx=20)
    checkVar = IntVar()
    advanced = Checkbutton(root_window, text="Advanced Mode", variable=checkVar)
    advanced.grid(row=4, column=1, pady=20, padx=20)
    convert_label = Label(root_window, text="2. Press the create Button.")
    convert_label.grid(row=3, column=0, pady=20, padx=20)
    convert_button = Button(root_window, text="create", command=lambda: check_and_create(checkVar.get()))
    convert_button.grid(row=3, column=1, pady=20, padx=20)
    root_window.mainloop()
    return None


def main():
    display_window()
    exit(0)


if __name__ == '__main__':
    main()
