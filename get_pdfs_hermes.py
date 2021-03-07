from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup as Bs
import time
import requests


import os

import PyPDF2
import re
#import pdftables
import numpy as np
import tabula
import pandas as pd

import matplotlib.pyplot as plt
#list_pdf =[]


def string_in_string(mot, liste):
    found = False
    for element in liste:
        if mot == element:
            found=True
            break
    return found

def download_pdfs():
    print("********** dowload_pdfs() **********")
    DRIVER_PATH = "C:\\Users\\virgi\\Downloads\\setups\\chromedriver.exe"
    options = Options()
    options.headless = True

    #browser = webdriver.Chrome(DRIVER_PATH)
    browser = webdriver.Chrome(options=options, executable_path=DRIVER_PATH)

    url = 'https://finance.hermes.com/en/publications/'
    # On lance Chrome
    browser.get(url)

    #*************************************************************************************************
    xpath_p1 =  '//*[@id="list-results"]/div['
    xpath_p2 = 11
    xpath_p3 = ']/div/button/label'

    '''
    //*[@id="listOfResults"]/div[1]/div[1]/button/div[1]

    //*[@id="list-results"]/div[71]/div/button/label
    '''
    #*************************************************************************************************
    # Fermeture du panneau des cookies, ouverture de "Filtrer" puis selection de "Rapport financier"
    #*************************************************************************************************
    cookie_close = '//*[@id="cookieContentClose"]'
    filtrer = '//*[@id="listOfResults"]/div[1]/div[1]/button/div[1]'
    financial_report = '//*[@id="listOfResults"]/div[1]/div[2]/ul/div[2]/ul/li[4]/label/div'

    list_tasks = [cookie_close, filtrer, financial_report]
    index_lt = 0

    while index_lt <=2 :
        try:
            xpath = list_tasks[index_lt]
            index_lt += 1
            loadButton = browser.find_element_by_xpath(xpath)
            
            time.sleep(2)
            loadButton.click()
            time.sleep(5)
            print("done \n")
        except Exception as e:
            print("***************** ya un soucis **************")
            print(e)
        
    #*************************************************************************************************
    # Clic "Afficher plus de documents" jusqu'à ce qu'il y ai une erreur (= tous les documents affichés)
    #*************************************************************************************************
    while True:
        try:
            xpath = xpath_p1+str(xpath_p2)+xpath_p3
            loadMoreButton = browser.find_element_by_xpath(xpath)
            int(xpath_p2)
            xpath_p2 += 10
            time.sleep(2)
            loadMoreButton.click()
            time.sleep(5)
            print("done \n")
        except Exception as e:
            print("***************** Il n'y a plus de document supplémentaire **************")
            print(e)
            break



    #*************************************************************************************************
    # Téléchargement des fichiers lorsqu'ils sont tous présents sur la page
    #*************************************************************************************************
    response = requests.get(url)

    time.sleep(10)
    soup = Bs(browser.page_source, 'html.parser')

    li=soup.find_all('a', {'class': 'document'})
    for i in li:
        print(i, '\n')

    pdf_links = [i['href']  for i in li if i['href'].endswith('pdf')]

    print("longueur pdf_links : ", len(pdf_links))
    in_li=0
    for link in pdf_links:

        in_li +=1
        print("link numero " + str(in_li) + " : " + link)

        pdf_string_index = link.find('pdf_file')
        #point_pdf_index = link.find('.pdf')
        if link[pdf_string_index + 17] != 'h':
            title_index = pdf_string_index + 28
        else:
            title_index = pdf_string_index + 17
        titre = link[title_index:]
        #list_pdf.append(titre)
        print("TITRE RETENU : " + titre + "\n")
        l_pdf = os.listdir(os.getcwd() + "/data/hermes_data")

        if not string_in_string(titre, l_pdf):
            r = requests.get(link, stream=True)
            with open('C:\\Users\\virgi\\Desktop\\virgile_stuff\\prog\\banking analyst\\financial_data\\data\\hermes_data\\' + titre, 'wb') as f:
                f.write(r.content)




#print(list_pdf)


def read_pdf_download_csv(pdf_name):
    print("********** read_pdf_download_csv('" + pdf_name + "') **********")
    # creating a pdf file object (on ouvre le fichier en code binaire)
    pdfFileObj = open(hermes_pdf_dir + "/" + pdf_name, 'rb') 

    # creating a pdf reader object  (Here, we create an object of PdfFileReader class of PyPDF2 module and  pass the pdf file object & get a pdf reader object)
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
    nbPages = pdfReader.numPages

    strings = strings_hermes[:]
    pages_id = []
    pages_id_titres = []
    # extract text and do the search
    for i in range(1, nbPages):
        pageObj = pdfReader.getPage(i)
        #print("this is page " + str(i)) 
        text = pageObj.extractText() 
        # print(Text)
        for string in strings :
            resSearch = re.search(string, text)
            if resSearch != None :
                pages_id.append(i+1)
                pages_id_titres.append((i+1,string))
        #print(resSearch)
    if pages_id == []:
        print("Aucune correspondance n'a été trouvée")
        return


#****************************************************************
    #1
    """
    tables = tabula.read_pdf(hermes_pdf_dir + "/" + pdf_name, pages = pages_id)
    folder_name = "tableaux"
    if not os.path.isdir(hermes_pdf_dir + "/" + folder_name):
        os.mkdir(hermes_pdf_dir + "/" + folder_name)

    #print(tables)
    print(pages_id)
    for (i, table) in enumerate(tables, start=1):
        table.to_excel(os.path.join(hermes_pdf_dir + "/" +folder_name, f"table_{i}.xlsx"), index=False)
        #table.to_excel(os.path.join(hermes_pdf_dir + "/" +folder_name, title.xlsx"), index=False)
"""

#****************************************************************
    #2
   
    if pages_id == []:
        print("Aucune correspondance n'a été trouvée")
        return

    dot_pdf = pdf_name.find(".pdf")
    #tableau_name = "tableaux"
    if not os.path.isdir(hermes_pdf_dir + "/tableaux"):
        os.mkdir(hermes_pdf_dir + "/tableaux")

    if not os.path.isdir(hermes_pdf_dir + "/tableaux/" + pdf_name[:dot_pdf]):
        os.mkdir(hermes_pdf_dir + "/tableaux/" + pdf_name[:dot_pdf])

    folder_name = hermes_pdf_dir + "/tableaux/" + pdf_name[:dot_pdf]

    for it_page,page in enumerate(pages_id, start=0):
        print("page : ", page)
        tables = tabula.read_pdf(hermes_pdf_dir + "/" + pdf_name, pages = page)
        
        
        for i,table in enumerate(tables, start=1):
            print( "titre retenu : " + pages_id_titres[it_page][1]  + " p" + str(page) + f"_{i}.xlsx")
            table.to_excel(os.path.join(folder_name, pages_id_titres[it_page][1]  + " p" + str(page) + f"_{i}.xlsx"), index=False)
    #problème les titres très souvent avec _1 à voir



    """
    #2)
        for i in range(1, nbPages):
        pageObj = pdfReader.getPage(i)
        #print("this is page " + str(i)) 
        text = pageObj.extractText() 
        # print(Text)
        for string in strings :
            resSearch = re.search(string, text)
            if resSearch != None :
                pages_id.append(i+1)
                pages_id_titres.append((i+1,string))
        #print(resSearch)
    if pages_id == []:
        print("Aucune correspondance n'a été trouvée")
        return

    folder_name = "tableaux"
    if not os.path.isdir(hermes_pdf_dir + "/" + folder_name):
        os.mkdir(hermes_pdf_dir + "/" + folder_name)

    for page in pages_id:
        tables = tabula.read_pdf(hermes_pdf_dir + "/" + pdf_name, pages = page)
        i=0
        for table in tables:
            i+=1
            table.to_excel(os.path.join(hermes_pdf_dir + "/" +str(folder_name),  str(pages_id_titres[page][1]) + f"_{i}.xlsx"), index=False)

#*******************************************************************************

    #1)
        for page in pages_id:
        table = tabula.read_pdf(hermes_pdf_dir + "/" + pdf_name, pages = page_id)
    folder_name = "tableaux"
    if not os.path.isdir(hermes_pdf_dir + "/" + folder_name):
        os.mkdir(hermes_pdf_dir + "/" + folder_name)

    #print(tables)
    print(pages_id)
    for (i, title) in pages_id_titres:
        #table.to_excel(os.path.join(hermes_pdf_dir + "/" +folder_name, f"table_{i}.xlsx"), index=False)
        table.to_excel(os.path.join(hermes_pdf_dir + "/" +folder_name, title.xlsx"), index=False)
        """




#*************************************************************************************************
# Suppression des fichiers sans tableau intéressant
# ATTENTION : changer méthode en fonction de ce qu'il faut supprimer OU PAS
# (pour l'instant qu'un tableau, soit type de données conservé)
#*************************************************************************************************
strings_hermes2 = ["In millions of euros", "Leater Goods & Saddlery", "Europe", "Number of shares as at 31 December", 
                "(millions of euros)", "millions of euros", "CASH FLOWS LINKED TO OPERATING ACTIVITIES", "Goodwill", "Equity"]
                #voir hermes2014 rapport annuel les pages : 303, 336, 362
def delete_bad_csv(folder):
    
    list_csv = os.listdir(folder)
    print("LIST CSV : ",list_csv)
    for csv_name in list_csv:
        print("********** delete_bad_csv(" + folder + "/" + csv_name + ") **********")
        found = False
        table = pd.read_excel(folder + "/" + csv_name) # j'utilise read_excel car .xlsx
        #print(table.columns)
        #print(table.columns[0],0)
        first_cell = table.columns[0]
        print("First cell : " + first_cell)
        for i in range(len(strings_hermes2)):
            string_index = first_cell.find(strings_hermes2[i])
            if first_cell=="Unnamed: 0":
                check = 2
                if len(table["Unnamed: 0"])>0:
                    while table["Unnamed: 0"][check]=="" and check < 15:
                        check +=1
                    cell = str(table["Unnamed: 0"][check])
                    string_index = cell.find(strings_hermes2[i])
            if string_index != -1:
                found = True
                print("found")
                break
        if found ==False:
            os.remove(folder + "/" + csv_name)
    

#************************************************ APPEL DES METHODES : *******************************************************
#download_pdfs()

actual_directory = os.getcwd()
hermes_pdf_dir = actual_directory + "/data/hermes_data"
list_pdf = os.listdir(hermes_pdf_dir)
list_pdf.pop(list_pdf.index('tableaux'))

strings_hermes = ["KEY CONSOLIDATED DATA", "ACTIVITY BY MÉTIER", "ACTIVITY BY GEOGRAPHIC AREA", 
            "KEY STOCK MARKET DATA", "SUMMARY CONSOLIDATED FINANCIAL STATEMENTS", 
            "CONSOLIDATED STATEMENT OF FINANCIAL POSITION", "EQUITY AND LIABILITIES",
            "CONSOLIDATED STATEMENT OF CASH FLOWS"]
    
#print("list_pdf : ", list_pdf)


#for pdf_name2 in list_pdf:
for i in range(14, len(list_pdf)):
    pdf_name2 = list_pdf[i]
    print("pdf_name : " + pdf_name2)
    if pdf_name2 != "hermes_2014_rapportannuel_en.pdf":
        #read_pdf_download_csv(pdf_name2)
        print("random shit")

csv_folder = actual_directory + "/data/hermes_data/tableaux"
list_hermes_folders_tables = os.listdir(csv_folder)
print("list_hermes_folders_tables : ", list_hermes_folders_tables)
for folder  in list_hermes_folders_tables:
    #delete_bad_csv(csv_folder + "/" + folder)
    print("random shit")
#généraliser à plusieurs pdf
#avoir 2010 -> 2020
#plot
#************************************************ FIN APPEL DES METHODES *******************************************************


import openpyxl
from pathlib import Path


def reverse(lst): 
    return [ele for ele in reversed(lst)] 

chemin_plot = actual_directory + "/data/hermes_data/tableaux/hermesinternational-urd-2019-en/KEY CONSOLIDATED DATA p18_1.xlsx"
def plot(fichier):
    years = []
    xlsx_file = Path(actual_directory + "/data/hermes_data/tableaux/hermesinternational-urd-2019-en", 'KEY CONSOLIDATED DATA p18_1.xlsx')
    wb_obj = openpyxl.load_workbook(xlsx_file)
    sheet = wb_obj.active


    """
    for row in sheet.iter_rows(max_row=6):
        for cell in row:
            print(cell.value, end=" ")
        print()
    """
    alphabet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    title = sheet['A2'].value

    cellx = sheet['B1']
    x = []
    column_number = 1
    while cellx.value!=None:
        x.append(cellx.value)
        column_number +=1
        cellx = sheet[alphabet[column_number] + str(1)]
 
    x = reverse(x)
    print("voila x : ", x)
    celly = sheet['B2']
    y = []
    column_number = 1
    while celly.value!=None:
        y.append(celly.value)
        column_number +=1
        celly = sheet[alphabet[column_number] + str(2)]
    y = reverse(y)
    print("voila y : ", y)
    plt.plot(x, y, color ='tab:blue')
    plt.title(title)
    plt.show()

plot(chemin_plot)