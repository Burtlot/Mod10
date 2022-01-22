from ast import Num
import email
from fileinput import filename
from multiprocessing import dummy
import numbers
from sqlite3 import Date
import sys
from tokenize import Number
from colorama import Fore, init, Back, Style
import openpyxl
from openpyxl.styles import NamedStyle
import re
from codicefiscale import codicefiscale


def hello (nome) :
    print(Back.BLUE,Fore.WHITE,'Leggo il file:',nome,Back.RESET+'\n')

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
    

regex = re.compile(r'([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+')

def isValidMail(emailaddress):
    if re.fullmatch(regex, emailaddress):
      return True
    else:
      return False

def checksesso (name) :
    if(name == 'M' or name == 'F') :
        return True
    else : 
        return False
    
def checkStudio (name) :
    if name == 'LM':
        return True
    elif name == 'IP':
        return True
    elif name == 'DM':
        return True
    elif name == 'IS':
        return True
    elif name == 'ME':
        return True
    elif name == 'EA':
        return True
    else:
        return False
        
def checkindirizzo (nome) :
    nome = checknome(nome)
    if nome.find("VIA") >= 0 : return True
    if nome.find("VIALE") >= 0 : return True
    if nome.find("LARGO") >= 0 : return True
    if nome.find("PIAZZA") >= 0 : return True
    return False
    
def checktelefono (nome) :
    nome = nome.replace("+39","")
    nome = nome.replace("+", "")
    nome = nome.replace(".", "")
    nome = nome.replace(" ", "")
    nome = nome.replace("-", "")
    nome = nome.replace("_", "")
    return nome

def checkvoto (voto):
    if 84 <= voto <= 140:
        return True
    else: 
        return False

errore = False
errdescr = ''
#init(convert=True)
print("\n")

#path = input("Inserisci il nome del file xls, ad ex- 70304.xlsx : ")

namefile = input("Inserisci il nome del file xls, ad ex- 70304.xlsx : ")

num_ext = namefile.find('.xlsx')
#print('result = '+str(num_ext))
if (num_ext < 0) : namefile = namefile + '.xlsx'

#input_col_name = input("Enter colname, ex- Endpoint : ")
try:
    #hello(path.rstrip)
    
    
    print(Fore.RESET)
    #path = "C:\\employee.xlsx"
    #wb_obj = openpyxl.load_workbook(path.strip())
    wb_obj = openpyxl.load_workbook(namefile)
    
    #leggi i nomi dei fogli presenti nel file excel
    sheetname = wb_obj.sheetnames
    #for i in range(0, len(sheetname)):
    #    print(Fore.BLUE + sheetname[i])

    #sheet = wb_obj.get_sheet_by_name(sheetname[0])
    sheet = wb_obj[sheetname[0]]
    max_column=sheet.max_column
    max_row=sheet.max_row
    
    hello(namefile+' - Colonne: '+str(max_column)+" / "+'Righe: '+str(max_row))
    #print(Fore.GREEN + 'Colonne: '+str(max_column))
    #print(Fore.RED + 'Righe: '+str(max_row))
    
    for i in range(2, max_row+1):
        errore = False
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
            errdescr = errdescr + riga + "->Errore Provincia Nascita: " + prvnascita + "\n "
        
        datanascita = sheet['E'+str(i)].value
        celldate = sheet['E'+str(i)]
        nsddmmyyyy=NamedStyle(name="cd"+str(i), number_format="DD/MM/YYYY")
        celldate.style = nsddmmyyyy
        
        #print(Fore.LIGHTGREEN_EX,celldate.value,Fore.RESET)
        
        indirizzo = sheet['F'+str(i)].value
        if not checkindirizzo(indirizzo) : 
            errore = True
            errdescr = errdescr + riga + "->Errore Controllare Indirizzo: " + indirizzo + "\n "
        
        cap = str(sheet['G'+str(i)].value)
        if not charexactly(str(cap),5) : 
            errore = True
            errdescr = errdescr + riga + "->Errore CAP: " + cap + "\n "
            
        luogoresidenza = sheet['H'+str(i)].value
        sheet['H'+str(i)] = checknome(luogoresidenza)
        
        provincia = sheet['I'+str(i)].value
        if not charexactly(provincia,2) : 
            errore = True
            errdescr = errdescr + riga + "->Errore Provincia Residenza: " + provincia + "\n "
        
        teldummy = sheet['J'+str(i)].value
        sheet['J'+str(i)] = ""
        
        cellulare = str(sheet['K'+str(i)].value)
        sheet['K'+str(i)] = checktelefono(cellulare)
        
        
        emailaddress = sheet['L'+str(i)].value
        if not isValidMail(emailaddress) :
            errore = True
            errdescr = errdescr + riga + "->Errore Indirizzo EMail: " + emailaddress + "\n "
            
        sesso = sheet['O'+str(i)].value
        if not checksesso(sesso) :
            errore = True
            errdescr = errdescr + riga + "->Errore Sesso: " + sesso + "\n "
        
        CodiceFIscaleFile = sheet['M'+str(i)].value
        CFCalcolate = codicefiscale.encode(surname=cognome, name=nome, sex=sesso, birthdate=datanascita, birthplace=luogonascita)
        if(CodiceFIscaleFile != CFCalcolate) : 
            errore = True
            errdescr = errdescr + riga + "->Controllare Codice Fiscale: " + CodiceFIscaleFile + " il calcolo a me risulta: " + CFCalcolate + "\n "
        
        titolostudio = sheet['N'+str(i)].value
        if not checkStudio(titolostudio) :
            errore = True
            errdescr = errdescr + riga + "->Errore Titolo di Studio: " + titolostudio + "\n "
        
        voto = sheet['P'+str(i)].value
        if not checkvoto(voto) :
            errore = True
            errdescr = errdescr + riga + "->Controllare Voto: " + str(voto) + "\n "
        
        
        #print(Fore.BLUE,CFCalcolate,CodiceFIscaleFile)
        
        #print(Fore.GREEN+riga+Fore.RED+"->"+Fore.CYAN+cognome,nome,luogonascita,prvnascita,datanascita,indirizzo,cap,luogoresidenza,provincia,cellulare,emailaddress,codicefiscale,titolostudio,sesso,voto)
        
    #for rowOfCellObject in sheet['A2':'P'+str(max_row)] :
    #    for cellObj in rowOfCellObject :
    #        print (Fore.YELLOW+cellObj.coordinate, cellObj.value)
    
    #for j in range(2, 5):
    #    salary_cell=sheet_obj.cell(row=j,column=2)
    #    if salary_cell.value > 1500 :
    #        salary_cell.value =  salary_cell.value+500
    print(Back.RED,Fore.YELLOW,errdescr,Back.RESET,Fore.RESET)
    
    savefile = namefile.replace(".xlsx", "_checked.xlsx")
    wb_obj.save(savefile)
except Exception as e:
    print(e)
    #print (Fore.RED + "Error : The file does not found")
print(Fore.GREEN + "###################### Ho riscritto il file "+savefile+" corregendo i campi testo ##############################")
