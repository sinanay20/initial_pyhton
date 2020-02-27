#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Name: icon_scripts
# Author: Samuel Bramm
# Co-Author: Benjamin Schmelcher
# Date: 2019-01-10
# Description: Collection of different scripts for ICON
# Version: v1.3.2
# Changelog
# v1.3.2 Fixed Encoding Bug in user_rights.py
# v1.3.1 Added gems_remove_tenants script
# v1.3 Eliminated DB data in config files, changed scripts to use db_connection class
# v1.2.1 Corrected several bugs in different scripts
# v1.2 Completely reworked compare_invoice_runs script for MBBras
# v1.1.1b Integrated iquote gems mapping scripts
# v1.0b Integrated multiple GEMS scripts
# v0.9 Integrated GEMS user right scripts
# v0.8b Integrated invoice runs comparison (beta)
# v0.7 Integrated migrated data tests and cofico feedback splitter
# v0.6 Integrated Contingent Mapping and doos scripts
# v0.5 Integrate user comparison script and updated rule config script
# v0.4 Integrate user mapping script
# v0.3 Integrate I23 parser
# v0.2 Integrate credit_notes.py and rule_config
# v0.1 Initial version

from Tkinter import *
from ttk import *
import os
import time
import subprocess


version = "v1.3.1"
root_window = Tk()
known_scripts = []
logfile = "./log/error.log"
scriptdir = "./script"

pages = []
current_page = 0
currentPageString = StringVar()



def exec_script(name):
    global scriptdir
    global root_window
    root_window.withdraw()
    try:
        subprocess.call([os.getcwd().replace('\\', '/') + "/Python-2.7.13/python.exe", scriptdir + "/" + name],
                        shell=True)
    except:
        log_writer("Error", "Can not execute script " + name + ".")
    root_window.deiconify()
    return None


def log_writer(type, message):
    writer = open("./log/" + logfile, "a")
    writer.write(time.strftime("%Y%m%d-%H%M%S ") + type + " " + message + "\n")
    writer.close()
    return None

def next_page():
    global current_page

    pages[current_page].grid_remove()

    if(current_page == len(pages)-1):
        current_page = 0
    else:
        current_page += 1

    pages[current_page].grid()
    currentPageString.set(str(current_page+1)+"/" + str(len(pages)))




def previous_page():
    global current_page

    pages[current_page].grid_remove()

    if(current_page == 0):
        current_page = len(pages)-1
    else:
        current_page -= 1
    pages[current_page].grid()
    currentPageString.set(str(current_page+1)+"/"+str(len(pages)))

def create_button_frame(root_window, container):
    buttonframe = Frame(root_window)
    previousButton = Button(buttonframe, text="<", command=lambda: previous_page())
    previousButton.grid(row=0, column=0, padx=10, pady=10)
    nextButton = Button(buttonframe, text=">", command=lambda: next_page())
    nextButton.grid(row=0, column=1, padx=10, pady=10)
    PageLabel = Label(buttonframe, textvariable=currentPageString)
    PageLabel.grid(row=1, column=0, columnspan=2, padx=10, pady=10)
    return buttonframe


def display_main_window():
    global root_window
    global known_scripts
    global current_page
    global currentPageString



    root_window.title("iCON Script collection " + version)

    container = Frame(root_window)

    buttonframe = create_button_frame(root_window, container)
    buttonframe.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    script_count = 0
    page_counter = 0


    container.grid(row=1, column=0, padx=10, pady=10)

    pages.append(Frame(container))
    # pages[page_counter].grid()
    # pages[page_counter].place(in_=container, x=0, y=0, relwidth=1, relheight=1)



    maxheight=0
    maxwidth=0
    for entry in known_scripts:
        Label(pages[page_counter], text=entry[0]).grid(row=(script_count + 1), column=0, pady=5, padx=5, sticky="w")
        Label(pages[page_counter], text=entry[1]).grid(row=(script_count + 1), column=1, pady=5, padx=5, sticky="w")
        Label(pages[page_counter], text=entry[2], wraplength=300).grid(row=(script_count + 1), column=2, pady=5, padx=5, sticky="w")
        Button(pages[page_counter], text="start script", command=lambda x=entry[0]: exec_script(x)).grid(row=(script_count + 1), column=3,
                                                                                                  pady=5, padx=5)


        script_count += 1
        if(script_count % 10 == 0):
            page_counter += 1
            pages.append(Frame(container))
            # pages[page_counter].grid(row=0, column=0)

    current_page = 0
    currentPageString.set("1/"+str(len(pages)))
    pages[0].grid()








    root_window.mainloop()


    #
    # Label(root_window, text="Scriptname").grid(row=0, column=0, pady=10, padx=10)
    # Label(root_window, text="Version").grid(row=0, column=1, pady=10, padx=10)
    # Label(root_window, text="Description").grid(row=0, column=2, pady=10, padx=10)
    # k=1
    # i=0
    # for entry in known_scripts:
    #     Label(root_window, text=entry[0]).grid(row=(i + 1), column=0+(k-1)*4, pady=5, padx=5)
    #     Label(root_window, text=entry[1]).grid(row=(i + 1), column=1+(k-1)*4, pady=5, padx=5)
    #     Label(root_window, text=entry[2]).grid(row=(i + 1), column=2+(k-1)*4, pady=5, padx=5)
    #     Button(root_window, text="start script", command=lambda x=entry[0]: exec_script(x)).grid(row=(i + 1), column=3+(k-1)*4,
    #                                                                                              pady=5, padx=5)
    #     i=i+1
    #     if(i % 18 == 0):
    #         k=k+1
    #         i=0
    #         Label(root_window, text="Scriptname").grid(row=0, column=0+(k-1)*4, pady=10, padx=10)
    #         Label(root_window, text="Version").grid(row=0, column=1+(k-1)*4, pady=10, padx=10)
    #         Label(root_window, text="Description").grid(row=0, column=2+(k-1)*4, pady=10, padx=10)
    # root_window.mainloop()
    return None



def build_script_db():
    global known_scripts
    global logfile
    global scriptdir
    try:
        for file in os.listdir(scriptdir):
            description = ""
            version = ""
            if file.endswith(".py"):
                try:
                    with open(scriptdir + "/" + file, 'r') as scriptfile:
                        for line in scriptfile:
                            if "Description:" in line:
                                description = line[15:].replace('\n', ' ').replace('\r', '').strip()
                            if "Version:" in line:
                                version = line[11:].replace('\n', ' ').replace('\r', '').strip()
                except IOError:
                    log_writer("Error", "Access to file" + file + " denied.")
                known_scripts.append([file, version, description])
    except IOError:
        log_writer("Error", "Can not list directory " + scriptdir + ".")
    return None


def main():
    s = Style()
    print s.theme_names()

    build_script_db()
    display_main_window()
    exit(0)


if __name__ == '__main__':
    main()
''' The cake is a lie
            ,:/+/-
            /M/              .,-=;//;-
       .:/= ;MH/,    ,=/+%$XH@MM#@:
      -$##@+$###@H@MMM#######H:.    -/H#
 .,H@H@ X######@ -H#####@+-     -+H###@X
  .,@##H;      +XM##M/,     =%@###@X;-
X%-  :M##########$.    .:%M###@%:
M##H,   +H@@@$/-.  ,;$M###@%,          -
M####M=,,---,.-%%H####M$:          ,+@##
@##################@/.         :%H##@$-
M###############H,         ;HM##M$=
#################.    .=$M##M$=
################H..;XM##M$=          .:+
M###################@%=           =+@MH%
@#################M/.         =+H#X%=
=+M###############M,      ,/X#H+:,
  .;XM###########H=   ,/X#H+:;
     .=+HM#######M+/+HM@+=.
         ,:/%XM####H/.
              ,.:=-.
'''