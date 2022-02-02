#!/usr/bin python3
from ast import Num
import email
from fileinput import filename
from multiprocessing import dummy
import numbers
from sqlite3 import Date
import sys
from tokenize import Number
from turtle import clear
from colorama import Fore, init, Back, Style
import openpyxl
from openpyxl.styles import NamedStyle
import re
from codicefiscale import codicefiscale
import requests
from requests.exceptions import HTTPError
from setuptools import setup
from datetime import datetime


__version__ = "0.1.9 del "+str(datetime.today().strftime('%d-%m-%Y'))
__annotations__="Controllo i dati prima di inviarli al Settore Tecnico FIGC"
__package__= "Corretto errore di eliminazione apice"


def release () :
    # Visualizzo la release alla partenza
    #print(Back.BLUE,Fore.WHITE,"Mod10 Rel."+__version__+" - "+__annotations__,Back.RESET+'\n')
    print(Back.BLUE,Fore.WHITE,"Mod10 Rel."+__version__,Back.RESET+'\n')

def hello (nome) :
    # Comunico quale file sto leggendo il numero di righe e colonne
    #print(Back.BLUE,Fore.WHITE,'Leggo il file:',nome,Back.RESET+'\n')
    print(Fore.GREEN,'# Leggo il file:',nome,Back.RESET+' #')
    
def is_empty(a):
    # Ritorna True se è vuoto
    return a == set()

def checknome (nome) :
    # Ritorno la stringa pulita da caratteri che non ci devono essere
    nome = nome.replace("'"," ")
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
    return nome.upper()

def charexactly (nome, numchar) :
    # Controllo che la stringa passata abbia l'esatto numero di caratteri ad esempio CAP di 5 chr
    conta = len(nome)
    if (conta > numchar or conta < numchar) :
        return False
    return True
    

regex = re.compile(r'([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+')

def isValidMail(emailaddress):
    # controllo che sia un indirizzo email scritto correttamente
    if re.fullmatch(regex, emailaddress):
      return True
    else:
      return False

def checksesso (name) :
    # Controllo cosa è inserito nel campo sesso
    if(name == 'M' or name == 'F') :
        return True
    else : 
        return False
    
def checkStudio (name) :
    # Controlllo che il titoli di studio sia tra quelli previsti dal modulo
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
    elif name == 'AC':
        return True
    elif name == 'EA':
        return True
    else:
        return False
        
def checkindirizzo (nome) :
    # Controllo che il campo indirizzo contenga un indicazione corretta
    nome = checknome(nome)
    if nome.find("VIA") >= 0 : return True
    if nome.find("VIALE") >= 0 : return True
    if nome.find("LARGO") >= 0 : return True
    if nome.find("PIAZZA") >= 0 : return True
    if nome.find("CORSO") >= 0 : return True
    if nome.find("CONTRADA") >= 0 : return True
    if nome.find("VICOLO") >= 0 : return True
    if nome.find("LOCALITA") >= 0 : return True
    
    return False
    
def checktelefono (nome) :
    # elimino i campi in esubero dal numero di cellulare
    nome = nome.replace("+39","")
    nome = nome.replace("+", "")
    nome = nome.replace(".", "")
    nome = nome.replace(" ", "")
    nome = nome.replace("-", "")
    nome = nome.replace("_", "")
    return nome

def checkvoto (voto):
    # Controllo che il valore del voto porti alla promozione altrimenti lo segnalo
    if 84 <= voto <= 140:
        return True
    else: 
        return False


def checkcomunenascita(cod_catasto) :
    # controllo il comune di nascita risalendo dal codice del catasto del codice fiscale sopratutto per inserire la nazione in caso di stranieri
    try:
        parametri = {"token": "05052021", "cod": cod_catasto}
        r = requests.post("http://segreteria.assoallenatori.it/api/catasto.php", data=parametri)
        return r.text.lstrip().upper()
        #print(r.json())
    except HTTPError as http_err:
        print(f'HTTP error occurred: {http_err}')
    except Exception as err:
        print(f'Other error occurred: {err}')

errore = False
errdescr = ''
#init(convert=True)
print("\n")

#path = input("Inserisci il nome del file xls, ad ex- 70304.xlsx : ")
if len(sys.argv) < 2 :
    namefile = input("Inserisci il nome del file xls, ad ex- 70304.xlsx : ")
else:
    namefile = sys.argv[1];     

# Controllo che il nome abbia l'estensione xlsx altrimenti la aggiungo
num_ext = namefile.find('.xlsx')
#print('result = '+str(num_ext))
if (num_ext < 0) : namefile = namefile + '.xlsx'

# Creo il file di salvataggio dei dati
savefile = namefile.replace(".xlsx", "_checked.xlsx")
#input_col_name = input("Enter colname, ex- Endpoint : ")
try:
    release()
        
    print(Fore.RESET)
    #Apro il file Excel
    wb_obj = openpyxl.load_workbook(namefile)
    
    #leggi i nomi dei fogli presenti nel file excel
    sheetname = wb_obj.sheetnames
    #for i in range(0, len(sheetname)):
    #    print(Fore.BLUE + sheetname[i])

    #sheet = wb_obj.get_sheet_by_name(sheetname[0])
    sheet = wb_obj[sheetname[0]]
    # Conto Righe e Colonne
    max_column=sheet.max_column
    max_row=sheet.max_row
    
    hello(namefile+' - Colonne: '+str(max_column)+" / "+'Righe: '+str(max_row))
    #print(Fore.GREEN + 'Colonne: '+str(max_column))
    #print(Fore.RED + 'Righe: '+str(max_row))
    
    # Leggo riga per riga partendo dalla seconda lascio stare le intestazioni
    for i in range(2, max_row+1):
        errore = False
        riga = str(i)
        
        cognome = sheet['A'+str(i)].value
        if is_empty(cognome) : 
            errore = True
            errdescr = errdescr + riga + "->Cognome vuoto\n "
        else:
            sheet['A'+str(i)] = checknome(cognome)        
        
        nome = sheet['B'+str(i)].value
        if is_empty(nome) : 
            errore = True
            errdescr = errdescr + riga + "->Nome vuoto\n "
        else:
            sheet['B'+str(i)] = checknome(nome)
        
        luogonascita = sheet['C'+str(i)].value
        #sheet['C'+str(i)] = checknome(luogonascita)
        
        prvnascita = sheet['D'+str(i)].value
        if is_empty(prvnascita) : 
            errore = True
            errdescr = errdescr + riga + "->Provincia Nascita vuota\n "
        else:
            if not charexactly(prvnascita,2) : 
                errore = True
                errdescr = errdescr + riga + "->Errore Provincia Nascita: " + prvnascita + "\n "
        
        datanascita = sheet['E'+str(i)].value
        if is_empty(datanascita) : 
            errore = True
            errdescr = errdescr + riga + "->Data di Nascita vuota\n "
        else :
            celldate = sheet['E'+str(i)]
            #print(celldate.style)
            nsddmmyyyy=NamedStyle(name="ddmmaaaa-"+str(i), number_format="DD/MM/YYYY")
            if(celldate.style == "Normale"):
                celldate.style = nsddmmyyyy
        
        #print(Fore.LIGHTGREEN_EX,celldate.value,Fore.RESET)
        
        #controllo che l'indirizzo sia completo di VIA PIAZZA ecc. ecc. 
        indirizzo = sheet['F'+str(i)].value
        if is_empty(indirizzo) : 
            errore = True
            errdescr = errdescr + riga + "->Indirizzo vuoto\n "
        else:
            if not checkindirizzo(indirizzo) : 
                errore = True
                errdescr = errdescr + riga + "->Errore Controllare Indirizzo: " + indirizzo + "\n "
            else :
                #se l'indirizzo è corretto lo copio
                sheet['F'+str(i)] = checknome(indirizzo)
            
        cap = str(sheet['G'+str(i)].value)
        if is_empty(cap) : 
            errore = True
            errdescr = errdescr + riga + "->CAP vuoto\n "
        else:
            if not charexactly(str(cap),5) : 
                errore = True
                errdescr = errdescr + riga + "->Errore CAP: " + cap + "\n "
            
        luogoresidenza = sheet['H'+str(i)].value
        if is_empty(luogoresidenza) : 
            errore = True
            errdescr = errdescr + riga + "->Luogo Residenza vuoto\n "
        else :
            sheet['H'+str(i)] = checknome(luogoresidenza)
        
        
        provincia = sheet['I'+str(i)].value
        if is_empty(provincia) : 
            errore = True
            errdescr = errdescr + riga + "->Provincia vuota\n "
        else:
            if not charexactly(provincia,2) : 
                errore = True
                errdescr = errdescr + riga + "->Errore Provincia Residenza: " + provincia + "\n "
        
        # Azzero il telefono fisso che non serve più
        teldummy = sheet['J'+str(i)].value
        sheet['J'+str(i)] = ""
        
        cellulare = str(sheet['K'+str(i)].value)
        if is_empty(cellulare) : 
            errore = True
            errdescr = errdescr + riga + "->Cellulare non presente\n "
        else :
            sheet['K'+str(i)] = checktelefono(cellulare)
        
        
        emailaddress = sheet['L'+str(i)].value
        if is_empty(emailaddress) : 
            errore = True
            errdescr = errdescr + riga + "->Manca indirizzo EMail\n "
        else:
            if not isValidMail(emailaddress) :
                errore = True
                errdescr = errdescr + riga + "->Errore Indirizzo EMail: " + emailaddress + "\n "
            
        sesso = sheet['O'+str(i)].value
        if is_empty(sesso) : 
            errore = True
            errdescr = errdescr + riga + "->Manca sesso\n "
        else:
            if not checksesso(sesso) :
                errore = True
                errdescr = errdescr + riga + "->Errore Sesso: " + sesso + "\n "
        
        CodiceFIscaleFile = sheet['M'+str(i)].value
        if is_empty(CodiceFIscaleFile) : 
            errore = True
            errdescr = errdescr + riga + "->Manca il Codice Fiscale\n "
        else:
            CFValid = codicefiscale.is_valid(CodiceFIscaleFile)
            if(not CFValid) : 
                errore = True
                #errdescr = errdescr + riga + "->Controllare Codice Fiscale: " + str(CodiceFIscaleFile) + " il calcolo a me risulta: " + str(CFCalcolate) + "\n "
                errdescr = errdescr + riga + "->Controllare Codice Fiscale: " + str(CodiceFIscaleFile) + "\n "
            else:
                # prendo il codice catastale dal codice fiscale e lo passo alla procedura
                catasto = CodiceFIscaleFile[11:15]
                prov = checkcomunenascita(catasto)
                sheet['C'+str(i)] = checkcomunenascita(catasto)
                #print(prov)
            
        #try:
        #CFCalcolate = codicefiscale.encode(surname=cognome, name=nome, sex=sesso, birthdate=datanascita, birthplace=luogonascita)
        #except ValueError as e:
            #CFCalcolate = ""
            #pass
            #CFCalcolate = codicefiscale.decode(CodiceFIscaleFile)
            #print(CFCalcolate['birthplace'])
            #raise ValueError("[codicefiscale] {}".format(e))
            #errdescr = errdescr + riga + "->Errore Calcolo CF: " + CFCalcolate['birthplace'] + "\n "    
            
        titolostudio = sheet['N'+str(i)].value
        if is_empty(titolostudio) : 
            errore = True
            errdescr = errdescr + riga + "->Manca il Titolo di Studio\n "
        else:
            if not checkStudio(titolostudio) :
                errore = True
                errdescr = errdescr + riga + "->Errore Titolo di Studio: " + titolostudio + "\n "
        
        voto = sheet['P'+str(i)].value
        if is_empty(voto) : 
            errore = True
            errdescr = errdescr + riga + "->Manca il Voto Finale\n "
        else:
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
    
    # Visualizzo gli errori in caso ci siano
    print(Back.RED+Fore.YELLOW+errdescr+Back.RESET+Fore.RESET)
    
    wb_obj.save(savefile)
    print(Fore.GREEN + "# scritto file "+savefile+" #")
except Exception as e:
    exception_type, exception_object, exception_traceback = sys.exc_info()
    filename = exception_traceback.tb_frame.f_code.co_filename
    line_number = exception_traceback.tb_lineno
    print(namefile,"Excel Line: ",str(i))
    print("Exception type: ", exception_type)
    print("File Name: ", filename)
    print("Code Line: ", line_number)
    #print(e)
    #print (Fore.RED + "Error : The file does not found")

