#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Name:  iquote_gems_mapper
# Author: Samuel Bramm
# Co-Author: Benjamin Schmelcher
# Date: 2018-08-27
# Description: creates gems upload files for iquote users from aat
# Version: v0.9b

from Tkinter import *
from ttk import *
from tkFileDialog import askopenfilename
import tkMessageBox
import os
from openpyxl import *
import time
import csv
import yaml

root = Tk()
config_path = './config/iquote_gems_mapper_scoped.yml'
input_file_path = ""
role_mapping = []
complete_product_profile = []
product_profile = []
product_profile2 = []
select_done = IntVar()
errors = []


def splited_lines_generator(lines, n):
    # split lines in chunks of size n
    for i in xrange(0, len(lines), n):
        yield lines[i: i + n]


def create_gems_file(user_list):
    # create cvs files for GEMS import
    global input_file_path
    global errors
    out_list = []
    for user in user_list: #uid, roles, profile, orgs
        if user[2] is not None:
            for product in product_profile:
                if user[2] == product[1]:
                    out_list.append([user[0], 'IQUOTE_PRODUCTPROFILE', '', user[2], '', ''])
            for product in product_profile2:
                if user[2] == product[1]:
                    out_list.append([user[0], 'IQUOTE_PRODUCTPROFILE2', '', user[2], '', ''])

        for role in user[1]:
            for org in user[3]:
                out_list.append([user[0], role, org, '', '', ''])
    for index, lines in enumerate(splited_lines_generator(out_list, 1999)):
        output_file_path = os.path.dirname(input_file_path) + "/gems_upload_{}_{}.csv".format(time.strftime("%Y_%m_%d"),
                                                                                              str(index).zfill(5))
        with open(output_file_path, 'wb') as csv_out_file:
            out_writer = csv.writer(csv_out_file, delimiter=";")
            out_writer.writerow(['UserID', 'RoleID', 'OrgScope', 'CustomScope', 'validfrom', 'validto'])
            for l in lines:
                out_writer.writerow(l)
    return os.path.dirname(input_file_path)

def decode_product_profile(app_parameters):
    # decode product_profile and map them to GEMS
    global complete_product_profile
    global errors
    profile_pattern = re.compile(r'IQUOTE.UserSpecificProductProfile = (.*)')
    for p in app_parameters.split(','):
        erg = re.search(profile_pattern, p)
        if erg:
            #print erg.group(1)
            map = [x[1] for x in complete_product_profile if x[0] == erg.group(1).strip()]
            if map:
                return map[0]
            else:
                errors.append("Error decoding product profile {}".format(p))
                return None

def decode_entitlements(entitlements):
    # decode entitlements from AAT and map them to GEMS
    roles = []
    global role_mapping
    global errors
    for e in entitlements.split(','):
        if e.upper().startswith('IQUOTE'):
            map = [x[1] for x in role_mapping if x[0] == e]
            if map:
                roles.append(map[0])
            else:
                errors.append("Error decoding entitlement {}. No mapping in config file found".format(e))
    return roles


def parse_iquote_file():
    # read iquote csv file and extract necessary data
    global input_file_path
    user_list = []
    with open(input_file_path, 'rb')as input_file:
        input_reader = csv.reader(input_file, delimiter=";")

        for i, line in enumerate(input_reader):
            roles = decode_entitlements(line[27])
            profile = decode_product_profile(line[28])
            orgs = []
            for entry in line[25].split(","):
                orgs.append(entry.strip())
            user_list.append([line[0], roles, profile, orgs])
    return user_list


def select_file():
    # read file path and store it in global variable
    global input_file_path
    global select_done
    input_file_path = askopenfilename(initialdir=os.getcwd(), title="Choose the iQUOTE file",
                                      filetypes=[("CSV", "*.csv")])
    select_done.set(1)
    return None


def read_config():
    # read mapping from configuration file and store it in global variable
    global config_path
    global role_mapping
    global complete_product_profile
    try:
        config = yaml.load(open(config_path))
        for role in config["entitlements"]:
            role_mapping.append([role, config["entitlements"][role]])
        for profile in config["product_profile"]:
            complete_product_profile.append([profile, config["product_profile"][profile]])
            product_profile.append([profile, config["product_profile"][profile]])
        for profile in config["product_profile2"]:
            complete_product_profile.append([profile, config["product_profile2"][profile]])
            product_profile2.append([profile, config["product_profile2"][profile]])

    except IOError:
        tkMessageBox.showerror("I/O Error", "Error opening configuration file at location {}".format(config_path))
    return None


def main():
    # display main menu and buttons
    global root
    global select_done
    global errors
    root.title("iQUOTE - GEMS Mapper")
    select_label = Label(root, text="1. Select the iQUOTE csv file.")
    select_label.grid(row=0, column=0, padx=10, pady=10)
    select_button = Button(root, text="select", command=lambda: select_file())
    select_button.grid(row=0, column=1, padx=10, pady=10)
    root.wait_variable(select_done)
    select_button.config(state=DISABLED)
    user_list = parse_iquote_file()
    out_file_path = create_gems_file(user_list)
    if len(errors) == 0:
        if tkMessageBox.askyesno("Success",
                                 "The gems files were successfully created in {}. Do you want to quit?".format(
                                     out_file_path)):
            exit(0)
    else:
        error_file_path = os.path.dirname(input_file_path) + "/errors_{}.csv".format(time.strftime("%Y_%m_%d"))
        with open(error_file_path, 'a') as error_file:
            for e in errors:
                error_file.write(e + "\n")
        tkMessageBox.showerror("Error",
                               "Errors occured during the creation of the GEMS file. Please correct the errors in {} and run the script again.".format(
                                   error_file_path))
    root.mainloop()
    exit(0)


if __name__ == '__main__':
    read_config()
    main()
