#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Name: gems_iquote_generator
# Author: Samuel Bramm
# Co-Author: Benjamin Schmelcher
# Date: 2018-09-06
# Description: add or delete iQUOTE roles for users in GEMS via cvs file
# Version: v0.9c

from Tkinter import *
from ttk import *
from tkFileDialog import askdirectory
import tkMessageBox
import os
from openpyxl import *
import time
import csv
import yaml

root = Tk()
config_path = './config/iquote_gems_mapper.yml'  # TODO change
input_file_path = ""
iquote_roles = []


def create_file(roles_update, orgs, users, action, p_profile):
    # build output list and write it to file
    user_list = [x.strip() for x in users.split(',')]
    role_list = [r[0] for r in roles_update if r[1].get() is True]
    out_list = []
    tkMessageBox.showinfo("Output folder", "Please specify the output folder for the generated files.")
    out_folder_path = askdirectory(initialdir=os.getcwd(), title="Choose the Output folder")
    for user in user_list:
        if action == "add":
            if p_profile is not None:
                out_list.append([user, "IQUOTE_PRODUCTPROFILE", '', p_profile, '', ''])
            if len(orgs) != 0:
                org_list = [x.strip() for x in orgs.split(',')]
                for org in org_list:
                    for role in role_list:
                        out_list.append([user, role, org, '', '', ''])
            else:
                for role in role_list:
                    out_list.append([user, role, '', '', '', ''])
        else:
            for role in role_list:
                out_list.append([user, role, '', '', '', ''])
    for idx, lines in enumerate(splited_lines_generator(out_list, 1999)):
        out_file_path = out_folder_path + "/gems_upload_{}_{}_{}_{}.csv".format('_'.join(user_list), action,
                                                                                time.strftime("%Y_%m_%d"),
                                                                                str(idx).zfill(5))
        with open(out_file_path, 'ab') as csv_out_file:
            out_writer = csv.writer(csv_out_file, delimiter=";")
            out_writer.writerow(
                ['UserID', 'RoleID', 'OrgScope', 'CustomScope', 'validfrom', 'validto'])
            for l in lines:
                out_writer.writerow(l)
    if tkMessageBox.askyesno("Success",
                             "The files have been saved in {}. Do you want to quit?".format(out_folder_path)):
        exit(0)


def splited_lines_generator(lines, n):
    # split lines in chunks of size n
    for i in xrange(0, len(lines), n):
        yield lines[i: i + n]


def load_config():
    # load configuration from file
    global iquote_roles
    global config_path
    try:
        config = yaml.load(open(config_path))
        for role in config["entitlements"]:
            iquote_roles.append(config["entitlements"][role])
    except IOError:
        tkMessageBox.showerror("I/O Error", "The configuration file is not available!")

    return None


def main():
    # diplay main window and buttons
    global iquote_roles
    user = StringVar()
    action = StringVar()
    orgs = StringVar()
    p_profile = StringVar()
    roles_update = []
    root.title("GEMS iQUOTE Generator")
    name_label = Label(root, text="1. Enter the username (or comma separated list).")
    name_label.grid(row=0, column=0, columnspan=3)
    name_input = Entry(root, textvariable=user)
    name_input.grid(row=0, column=3)
    action_label = Label(root, text="2. Select the desired action.")
    action_label.grid(row=1, column=0, columnspan=3)
    action_menu = OptionMenu(root, action, "add", *sorted(["add", "delete"]))
    action_menu.grid(row=1, column=3)
    orgs_label = Label(root,
                       text="3. [optional] Specify the organisations (comma separated list) if action is \"add\".")
    orgs_label.grid(row=2, column=0, columnspan=3)
    orgs_entry = Entry(root, textvariable=orgs)
    orgs_entry.grid(row=2, column=3)
    c_scope_label = Label(root, text="4. [optional] Specify user specific product profile.")
    c_scope_label.grid(row=3, column=0, columnspan=3)
    c_scope_entry = Entry(root, textvariable=p_profile)
    c_scope_entry.grid(row=3, column=3)
    roles_label = Label(root, text="4. Select the specific roles you want to change.")
    roles_label.grid(row=4, column=0, columnspan=3)
    row = 5
    new_row = 5
    for idx, roles in enumerate(splited_lines_generator(iquote_roles, 3)):
        col = 0
        for i, r in enumerate(roles):
            u = BooleanVar()
            l = Label(root, text=r)
            l.grid(row=row + idx, column=col)
            c = Checkbutton(root, variable=u)
            c.grid(row=row + idx, column=col + 1)
            roles_update.append([r, u])
            col += 2
        new_row = row + idx

    create_label = Label(root, text="5. Press the create Button")
    create_label.grid(row=new_row + 1, column=0)
    create_button = Button(root, text="create",
                           command=lambda: create_file(roles_update, orgs.get(), user.get(), action.get(),
                                                       p_profile.get()))
    create_button.grid(row=new_row + 1, column=1)
    root.mainloop()
    exit(0)


if __name__ == '__main__':
    load_config()
    main()
