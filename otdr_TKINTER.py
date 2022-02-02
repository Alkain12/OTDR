# from glob import glob
import tkinter as tk
import tkinter.ttk as ttk
from tkinter.ttk import Progressbar
from tkinter import filedialog
import pyOTDR
import os
import re
from openpyxl import Workbook
import tkinter
import sys

root = tkinter.Tk()
root.title("Pliki Sor do XLSX")
root.geometry('480x480+100+100')

pliksor = re.compile(r'.sor')
fala = re.compile(r'1310')
drop = re.compile(r'drop')
wb = Workbook()
ws = wb.active
ws['A1'] = "Złącze numer"
ws['H1'] = "Złącze numer"
ws['B1'] = "[dB]"
ws['I1'] = "[dB]"
ws['C1'] = "[km]"
ws['J1'] = "[km]"
ws['F1'] = "Złącze A"
ws['M1'] = "Złącze A"

# baz_fold = input("podaj całą ściężkę\n")
baz_fold = str()
# nazwapliku = input("Podaj nazwę pliku końcowego\n")
nazwapliku = str()
licz = 0
liczpliki = 0
var = tk.IntVar()


def miejsce_button():
    def f():
        miejsce_butt: str = tkinter.filedialog.askdirectory()
        miejsce.insert(-1, miejsce_butt)

    return f()


def pobranie():
    def f():
        miejsce2 = miejsce.get()
        plik2 = plik.get()
        # print(miejsce2)
        # print(plik2)
        liczenie(miejsce2, plik2)

    return f()


def liczenie(baz_fold, nazwapliku):
    licz = 0
    licz2 = 2

    # okno_licz = tkinter.Toplevel(root)
    # okno_licz.geometry('200x200')
    # okno_licz.title("Postęp")
    #
    # pasek = Progressbar(okno_licz, length=100, style='black.Horizontal.TProgressbar')
    # pasek.pack(expand=True)
    # pasek['value'] = licz

    # licztdqm = 0
    sub_folders = []
    for dir, sub_dirs, files in os.walk(baz_fold):
        sub_folders.extend(sub_dirs)
        sub_folders1 = str(f"{dir}")
        # print(f'foldery {sub_folders1}')

        for plik in os.listdir(sub_folders1):
            if pliksor.search(plik):
                # print(f'nazwa pliku {plik}')
                if fala.search(plik):
                    new = pyOTDR.sorparse(str(sub_folders1) + '\\' + str(plik))
                    new2 = new[1]
                    nazwa2 = plik.split(sep='_')
                    # print(nazwa2)
                    # print(new2["GenParams"])
                    nazwa3 = nazwa2[1]
                    if nazwa3 == "DROP":
                        nazwa = new2["GenParams"]["location A"] + f' D' + nazwa2[2]
                    if nazwa3 == "OLT":
                        nazwa = new2["GenParams"]["location A"] + f' SP' + nazwa2[2]
                    if nazwa3 != "DROP" and nazwa3 != "OLT":
                        nazwa = new2["GenParams"]["location B"] + f' P' + nazwa2[2]
                    # print(nazwa2)
                    db = new2["KeyEvents"]['Summary']["total loss"]
                    lenght = new2["KeyEvents"]['Summary']["ORL finish"]
                    ws.cell(row=licz2, column=1).value = nazwa
                    # wavelenght = new2["GenParams"]["wavelength"]
                    ws.cell(row=licz2, column=2).value = db
                    ws.cell(row=licz2, column=3).value = lenght
                    if new2["KeyEvents"]['event 1']["refl loss"] != "0.000":
                        ws.cell(row=licz2, column=6).value = new2["KeyEvents"]['event 1']["refl loss"]
                    licz2 -= 1
                else:
                    new = pyOTDR.sorparse(str(sub_folders1) + '\\' + str(plik))
                    new2 = new[1]
                    nazwa2 = plik.split(sep='_')
                    # print(nazwa2)
                    nazwa3 = nazwa2[1]
                    if nazwa3 == "DROP":
                        nazwa = new2["GenParams"]["location A"] + f' D' + nazwa2[2]
                    if nazwa3 == "OLT":
                        nazwa = new2["GenParams"]["location A"] + f' SP' + nazwa2[2]
                    if nazwa3 != "DROP" and nazwa3 != "OLT":
                        nazwa = new2["GenParams"]["location B"] + f' P' + nazwa2[2]
                    # print(plik)
                    db = new2["KeyEvents"]['Summary']["total loss"]
                    lenght = new2["KeyEvents"]['Summary']["ORL finish"]
                    ws.cell(row=licz2, column=8).value = nazwa
                    # wavelenght = new2["GenParams"]["wavelength"]
                    ws.cell(row=licz2, column=9).value = db
                    ws.cell(row=licz2, column=10).value = lenght
                    if new2["KeyEvents"]['event 1']["refl loss"] != "0.000":
                        ws.cell(row=licz2, column=13).value = new2["KeyEvents"]['event 1']["refl loss"]
                licz += 1
                licz2 += 1
            # pasek['value'] = licz
            print(licz)

    wb.save(f"{baz_fold}\{nazwapliku}.xlsx")
    # print(licz)
    new_wind = tkinter.Toplevel(root)
    new_wind.title("Koniec pracy")
    tkinter.Label(new_wind, text="Skończone").pack()
    # sys.exit()


tkinter.Label(root, text='Gdzie są pliki?').grid(row=0, column=0, padx=2, pady=2)
b1 = tkinter.Button(root, text='Wskaż miejsce ', command=miejsce_button).grid(row=0, column=2, pady=2, padx=2)
miejsce = tkinter.Entry(root)
# b1 = tkinter.Button(root, text='Wskaż miejsce ', command=miejsce_button).grid(row=0, column=2,pady=2,padx=2)
miejsce.grid(row=0, column=1)

tkinter.Label(root, text='Nazwa pliku xlsx').grid(row=1, column=0, padx=2, pady=2)
plik = tkinter.Entry(root)
plik.grid(row=1, column=1)

b2 = tkinter.Button(root, text='Twórz XLSX, plik ', command=pobranie).grid(row=2)

style = ttk.Style()
style.theme_use('default')
style.configure("black.Horizontal.TProgressbar", background='brown')

pasek = Progressbar(root, length=00, style='black.Horizontal.TProgressbar')
pasek.grid(columnspan=3)
# pasek['value'] = licz

root.mainloop()
