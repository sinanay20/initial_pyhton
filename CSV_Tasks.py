import csv

# CSV-Datei mit einigen Firmen in Baden-Württemberg
# TODO: Selektion der Daten nach Firmenname, ID und Ansprechpartner
# @Return: 1. Ausgabe der Firmen nach Selektion in einer Liste und in Strings.
#          2. Anzahl der Firmen.
#          3. Ausgabe ohne Header.

f_id = "Firmen-ID:"
f_name = "Firma:"
f_ansprechpartner = "Ansprechpartner:"
space = ", "

print("----------SortedCompanysString---------")
with open("firmen.csv", 'r') as file:
    next(file)
    counter = 0
    for line in file:
        data = line.strip().split(";")
        sorted_companys = f_id+data[0]+space + f_name + \
            data[1]+space+f_ansprechpartner+data[11]
        counter += 1
        print(sorted_companys)
    print("Anzahl an Firmen: "+str(counter))


print("\n----------SortedCompanysList---------")
with open("firmen.csv", 'r') as file:
    next(file)
    counter = 0
    for line in file:
        data = line.strip().split(";")
        sorted_companys = f_id+data[0]+f_name + \
            data[1]+f_ansprechpartner+data[11]
        sorted_companys_list = sorted_companys.split()
        counter += 1
        print(sorted_companys_list)
    print("Anzahl an Firmen: "+str(counter))


# TODO: Hinzufügen von Werten in eine CSV-Datei
# @Return: Einträge in der CSV-Datei in einer Liste und in Strings.

print("\n----------CSV-Write-List---------")

vornamen = ["Sinan", "Alexander", "Laura", "Meltem"]
nachnamen = ["Ayten", "Böhm", "Eberhard", "Dökmedemir"]
semester = [str(6), str(6), str(7), str(6)]

with open('write_CSV2.csv', 'w', newline='') as new_file:
    felder = ["Vorname", "Nachname", "Semester"]
    w_csv = csv.DictWriter(new_file, fieldnames=felder)
    w_csv.writeheader()
    # Anlegen der Zeilen & Spalten
    w_csv.writerow(
        {"Vorname": vornamen[0], "Nachname": nachnamen[0], "Semester": semester[0]})
    w_csv.writerow(
        {"Vorname": vornamen[1], "Nachname": nachnamen[1], "Semester": semester[1]})
    w_csv.writerow(
        {"Vorname": vornamen[2], "Nachname": nachnamen[2], "Semester": semester[2]})
    w_csv.writerow(
        {"Vorname": vornamen[3], "Nachname": nachnamen[3], "Semester": semester[3]})

with open('write_CSV2.csv', 'r') as csv_file:
    for line in csv_file:
        open_csv = line.strip().split(";")
        print(open_csv)

print("\n----------CSV-Write-String---------")
str_vorname = '''Vorname:'''
str_nachname = '''Nachname:'''
str_semester = '''Semester:'''
space = ", "

with open('write_CSV2.csv', 'r') as csv_file:
    for line in csv_file:
        open_csv = str_vorname + \
            vornamen[0] + space + str_nachname+nachnamen[1] + \
            space + str_semester + space+semester[2]
        print(open_csv)
