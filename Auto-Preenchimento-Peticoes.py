import tkinter as tk
import klembord
import tkinter as tk
from tkinter import *
from tkinter import messagebox
from tkinter.font import BOLD
from tkinter.ttk import *
from tkinter import ttk
from tkinter.constants import CENTER, TOP, DISABLED
from openpyxl import load_workbook
import pyexcel as p
import pyexcel_xls
import pyexcel_xlsx
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import os
import time
import glob
import winshell
from win32com.client import Dispatch
from functools import partial
import datetime
from datetime import date
import bisect 

# oii
root = tk.Tk()
text = tk.Text(root)
text.pack(fill='both', expand=True)

text.tag_configure('normal', font='TimesNewRoman 12')
text.tag_configure('bold', font='TimesNewRoman 12 bold')
text.tag_configure('center', justify='center', font='TimesNewRoman 12 bold')
text.tag_configure('center2', justify='center', font='TimesNewRoman 12')
text.tag_configure('justified', justify='right', font='TimesNewRoman 12')
text.tag_configure('justified2', justify='right', font='TimesNewRoman 12 bold')

TAG_TO_HTML = {
    ('tagon', 'bold'): '<b>',
    ('tagoff', 'bold'): '</b>',
    ('tagon', 'center'): '<center>', 
    ('tagoff', 'center'): '</center>', 
    ('tagon', 'center2'): '<center>',
    ('tagoff', 'center2'): '</center>',
    ('tagon', 'justified'): '<p align="justify">',
    ('tagoff', 'justified'): '</p>',
    ('tagon', 'justified2'): '<p align="justify">',
    ('tagoff', 'justified2'): '</p>',
}

def copy_rich_text(event):
    try:
        txt = text.get('sel.first', 'sel.last')
    except tk.TclError:
        # no selection
        return "break"
    content = text.dump('sel.first', 'sel.last', tag=True, text=True)
    html_text = []
    for key, value, index in content:
        if key == "text":
            html_text.append(value)
        else:
            html_text.append(TAG_TO_HTML.get((key, value), ''))
    klembord.set_with_rich_text(txt, ''.join(html_text))
    return "break"  # prevent class binding to be triggered

text.bind('<Control-c>', copy_rich_text)


# ------------------------------------- SANTA CATARINA ----------------------------------

# load file, sheet
wb = load_workbook('processos_sc.xlsx')
ws = wb.active


# formatting excel sheet
for i in range(0,2):
    ws.delete_rows(1)

for i in range(0,3):
    ws.delete_cols(7)


# row number variable
max_rows = ws.max_row

# append list variables
processos = []
tipo = []
classe = []
autores = []
reus = []
cidade = []
assunto = []

# copy num processos
for i in range(1,max_rows+1):
    processos.append(ws.cell(row = i, column = 1).value)

# copy classe
for i in range(1,max_rows+1):
    tipo.append(ws.cell(row = i, column = 2).value)

# copy autores
for i in range(1,max_rows+1):
    autores.append(ws.cell(row = i, column = 3).value)

# copy reu(s)
for i in range(1,max_rows+1):
    reus.append(ws.cell(row = i, column = 4).value)

# copy cidade
for i in range(1,max_rows+1):
    cidade.append(ws.cell(row = i, column = 5).value)

# copy assunto 
for i in range(1,max_rows+1):
    assunto.append(ws.cell(row = i, column = 6).value)


# append lists
numero_id = []
numero_processo = []

# get numero_id
for i in range(0,max_rows):
    numero_id.append(processos[i])

numero_id = [x[21:25] for x in numero_id]



# get numero_processo
for i in range(0,max_rows):
    numero_processo.append(processos[i])

numero_processo = [x[:25] for x in numero_processo]

pesquisa_processo = ""


#pesquisa = input(pesquisa_processo)

#print(pesquisa)

#index = numero_processo.index(pesquisa)

index = 5
print(index)



estado = "SC"
vara = "1"
header = "EXMO. SR. DR. JUIZ DE DIREITO DA {}ª VARA CÍVEL DA COMARCA DE {}/{}.\n \n \n \n".format(vara, cidade[index].upper(), estado)
autos = "Autos nº {}\n \n".format(numero_processo[index])
reu_final = "Réu: {}\n \n".format(reus[index])
classe = "Ação: {}\n \n \n".format(tipo[index])
nossa_parte = "{}".format(autores[index])
texto = ", vem respeitosamente à presença de Vossa Excelência através de seu procurador que esta subscreve, expor e requerer o que segue: \n \n \n \n \n"


diadehoje = date.today().strftime("%d/%m/%Y")

dia = diadehoje[:2]
mes = diadehoje[3:5]
ano = diadehoje[8:]

meses_extenso = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
mes = int(mes)
mes = meses_extenso[mes-1]


data = "Nestes termos\nPede deferimento.\nMafra, {} de {} de 20{}.\n \n".format(dia, mes, ano)

nome = "MÁRCIO MAGNABOSCO DA SILVA\n OAB/SC 9738 – OAB/PR 20962"


text.insert("end",header,"center")
text.insert("end",autos,"bold")
text.insert("end",reu_final,"normal")
text.insert("end",classe,"normal")
text.insert("end",nossa_parte,"justified2")
text.insert("end",texto,"justified")
text.insert("end",data, "center2")
text.insert("end",nome,"center")

#text.insert("1.0", "Author et al. (2012). The title of the article. ")
#text.insert("end", "Journal Name", "italic")
#text.insert("end", ", ")
#text.insert("end", "2", "bold")
#text.insert("end", "(599), 1–5.")

root.mainloop()