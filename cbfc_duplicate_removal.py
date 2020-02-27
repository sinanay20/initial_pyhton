#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Name: CBFC Duplicate Removal
# Author: Benjamin Schmelcher
# Date: 20190128
# Description: Removes duplicate entries in multiple CBFC files
# Version: 0.0.1
import io
import os
import tkMessageBox
from Tkinter import Tk, IntVar
from tkFileDialog import askdirectory
from ttk import Label, Button

root = Tk()
directory_selected = IntVar()
files_created = IntVar()
cbfc_path = ""
out_path = ""


def select_directory():
    global cbfc_path
    cbfc_path = askdirectory(title="Choose directory containing CBFC files", initialdir="C:/Users/SCHMEBE/Dokumente/Tejmur/2019/cbfc file duplicate removal/cbFC_Fdb_20190118_153006_0008433376")

    directory_selected.set(1)


def edit_files():
    global out_path

    out_path = cbfc_path + "/output"
    try:
        os.stat(out_path)
        if tkMessageBox.askyesno("Output Folder already exists", "Output Folder already exists in the chosen directory.\nClear output folder?"):
            for file in os.listdir(out_path):
                if file.endswith(".txt"):
                    try:
                        os.remove(os.path.join(out_path, file))
                    except:
                        tkMessageBox.showerror("Error deleting file", "Could not delete file " + file + " in the output directory.\nPlease make sure to close all instances of it and restart the Script.")
                        exit(-1)

    except:
        os.mkdir(out_path)

    for filename in os.listdir(cbfc_path):
        temp_file = []
        filepath = cbfc_path+"/"+filename
        outpath = out_path+"/"+filename
        if filepath.endswith(".txt"):
            with io.open(filepath, 'r', encoding="utf-8") as cbfc:
                for idx, line in enumerate(cbfc):
                    if idx <=1:
                        temp_file.append(line)
                    else:
                        break
            if temp_file[0][39] != "1":
                print "Error in file: " + filepath

            with io.open(outpath, 'w', encoding="utf-8") as out:
                out.writelines(temp_file)

    files_created.set(1)



def main():
    global out_path
    select_directory_label = Label(root, text="1. Select directory containing all .txt CBFC files.")
    select_directory_label.grid(row=0, column=0, padx=10, pady=10)
    select_directory_button = Button(root, text="select", command=lambda: select_directory())
    select_directory_button.grid(row=0, column=1, padx=10, pady=10)
    root.wait_variable(directory_selected)

    edit_label = Label(root, text="2. Click the edit button to remove duplicates")
    edit_label.grid(row=1, column=0, padx=10, pady=10)
    edit_button = Button(root, text="edit", command=lambda: edit_files())
    edit_button.grid(row=1, column=1, padx=10, pady=10)
    root.wait_variable(files_created)



    tkMessageBox.showinfo("Success", "Succes.\nThe files were created in " + out_path + ".\nExiting script.")

    exit(0)
    root.mainloop()

if __name__ == '__main__':
    main()