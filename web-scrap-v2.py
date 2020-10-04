# -*- coding: utf-8 -*-
"""
Created on Thu Mar 28 23:29:33 2019

@author: Hugo Pfeffer
"""
import xlsxwriter
from bs4 import BeautifulSoup as soup
from urllib.request import urlopen as uReq

my_url = 'http://contadores.cnt.br/agenda-tributaria/estadual/distrito-federal/2019/05.html'

while True:
    try: 
        mes = int(input("Mês desejado: "))
        ano = int(input("Ano desejado: "))
        break
    except ValueError:
        print("Hmm, parece que alguém andou tentando quebrar meu código. Use apenas números na data!")

filename = xlsxwriter.Workbook("agenda-import.xlsx")
sheet = filename.add_worksheet("Agenda")

#Cria os headres no Excel
sheet.write(0, 0, "Subject")
sheet.write(0, 1, "Start Date")
sheet.write(0, 2, "Start Time")
sheet.write(0, 3, "End Date")
sheet.write(0, 4, "End Time")
sheet.write(0, 5, "All Day Event")
sheet.write(0, 6, "Description")
sheet.write(0, 7, "Location")
sheet.write(0, 8, "Private")

#================BeatifulSoup começa aqui! ================
#Requisita a URL
uClient = uReq(my_url)
#Baixa o conteudo da URL e fecha o pedido em seguida
page_html = uClient.read()
uClient.close()

#Analiza o HTML
page_soup = soup(page_html, "html.parser")

containers = page_soup.findAll("tbody", {"class":"tributos-do-dia"})
row_tit = 1
row_dia = 1
row_cont = 1
row_private = 1
row_allday = 1
row_loc = 1
row_enddate = 1
format2 = filename.add_format({'num_format':'mm/dd/yy'})

print("Criando Arquivo...")
print("Misturando sangue de galinha preta com cachaça...")
for container in containers:
    #Salva o texto em vars. 
    titulo_container = container.findAll("td", {"class":"titulo"})
    dia = str(mes) + "/" + container.td.contents[0] + "/" + str(ano)
    conteudo_container = container.findAll("td", {"class":"conteudo"})
    for titulos in titulo_container:
        titulo = titulos.getText()
        sheet.write(row_tit, 0, titulo.replace("- ", ""))
        sheet.write(row_dia, 1, dia, format2)
        sheet.write(row_enddate, 3, dia, format2)
        sheet.write(row_private, 8, "FALSE")
        sheet.write(row_allday, 5, "TRUE")
        sheet.write(row_loc, 7, "Brazil")
        row_private += 1
        row_allday += 1
        row_loc += 1
        row_dia += 1
        row_tit += 1
        row_enddate += 1
    for conteudos in conteudo_container:
        conteudo = conteudos.getText()
        sheet.write(row_cont, 6, conteudo)
        row_cont += 1
print("Pronto para importação. Lembre-se... You don't talk about the fight club")
filename.close()




