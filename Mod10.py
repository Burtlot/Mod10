from ast import Num
import email
from multiprocessing import dummy
import numbers
import sys
from colorama import Fore, init, Back, Style
import openpyxl
import re

def hello (nome) :
    print(Back.BLUE,Fore.WHITE,'Leggo il file:',nome)
    print(Back.RESET)

def checknome (nome) :
    nome = nome.replace("'","")
    nome = nome.replace("à", "a")
    nome = nome.replace("è", "e")
    nome = nome.replace("é", "e")
    nome = nome.replace("ì", "i")
    nome = nome.replace("ò", "o")
    nome = nome.replace("ù", "u")
    nome = nome.replace("À", "A")
    nome = nome.replace("È", "E")
    nome = nome.replace("Ì", "I")
    nome = nome.replace("Ò", "O")
    nome = nome.replace("Ù", "U")
    return nome

def charexactly (nome, numchar) :
    conta = len(nome)
    if (conta > numchar or conta < numchar) :
        return False
    return True
    

errore = False
errdescr = ''
#init(convert=True)
print("\n")
path = input("Inserisci il nome del file xls, ad ex- 70304.xlsx : ")
#input_col_name = input("Enter colname, ex- Endpoint : ")
try:
    hello(path.rstrip)
    print(Fore.RESET)
    #path = "C:\\employee.xlsx"
    wb_obj = openpyxl.load_workbook(path.strip())
    
    #leggi i nomi dei fogli presenti nel file excel
    sheetname = wb_obj.sheetnames
    #for i in range(0, len(sheetname)):
    #    print(Fore.BLUE + sheetname[i])

    sheet = wb_obj.get_sheet_by_name(sheetname[0])
    
    # from the active attribute 
    sheet_obj = wb_obj.active

    # get max column count
    max_column=sheet_obj.max_column
    max_row=sheet_obj.max_row
    
    print(Fore.GREEN + 'Colonne: '+str(max_column))
    print(Fore.RED + 'Righe: '+str(max_row))
    
    for i in range(2, max_row+1):
        riga = str(i)
        
        cognome = sheet['A'+str(i)].value
        sheet['A'+str(i)] = checknome(cognome)        
        
        nome = sheet['B'+str(i)].value
        sheet['B'+str(i)] = checknome(nome)
        
        luogonascita = sheet['C'+str(i)].value
        sheet['C'+str(i)] = checknome(luogonascita)
        
        prvnascita = sheet['D'+str(i)].value
        if not charexactly(prvnascita,2) : 
            errore = True
            errdescr = errdescr + Fore.RED + riga + "Errore Provincia" + "\n"
        
        datanascita = sheet['E'+str(i)].value
        indirizzo = sheet['F'+str(i)].value
        cap = str(sheet['G'+str(i)].value)
        luogoresidenza = sheet['H'+str(i)].value
        provincia = sheet['I'+str(i)].value
        teldummy = sheet['J'+str(i)].value
        cellulare = str(sheet['K'+str(i)].value)
        emailaddress = sheet['L'+str(i)].value
        codicefiscale = sheet['M'+str(i)].value
        titolostudio = sheet['N'+str(i)].value
        sesso = sheet['O'+str(i)].value
        voto = str(sheet['P'+str(i)].value)
        print(Fore.GREEN+riga+Fore.RED+"->"+Fore.CYAN+cognome,nome,luogonascita,prvnascita,datanascita,indirizzo,cap,luogoresidenza,provincia,cellulare,emailaddress,codicefiscale,titolostudio,sesso,voto)
        
    #for rowOfCellObject in sheet['A2':'P'+str(max_row)] :
    #    for cellObj in rowOfCellObject :
    #        print (Fore.YELLOW+cellObj.coordinate, cellObj.value)
    
    #for j in range(2, 5):
    #    salary_cell=sheet_obj.cell(row=j,column=2)
    #    if salary_cell.value > 1500 :
    #        salary_cell.value =  salary_cell.value+500

    wb_obj.save("visto.xlsx")
except Exception as e:
    print(e)
    #print (Fore.RED + "Error : The file does not found")
print(Fore.GREEN + "###################### Successfully! Excel file has been read/write. ##############################")
