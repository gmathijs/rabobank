#!/usr/bin/env python
# coding: utf-8

# # Rabobank Import CSV

# ### Beschrijving
# Dit is een universele file om grootboek en kleinboek codes alvast op maandbasis toe te voegen.
# Kan worden gebruikt voor een geheel jaar of in de loop van een jaar
# Inlezen van een of meerdere csv file van de Rabobank.
# 
# Manueel:clear
# Toevoegen van de juiste grootboek codes
# Toevoegen jaar en maand
# Verwijderen van kolommen die je toch niet nodig hebt
# Hernoemen van de kolommen naar iets begerijpelijks
# Creatie van uitvoer files voor verdere verwerking
# 
# #### Randvoorwaarde:
# Deze file gaat uit dat aan ieder volgnummer geen grootboek categorie is toegevoegd.
# Het originele CSV bestand van de rabobank blijft onaangetast
# 
# ### benodigde bestanden
# - csv download files van de rabobank met alle rekeningen over een periode bijvoorbeeld maandper maand. deze files starten met CSV_A of CSV_CC stop die in een onderliggend directorie van het jaartal bijvoorbeeld \2023 
# -       Meerdere csv bestanden zijn mogelijk, duplicaten worden eruit gegooid
# 
# #### Bestanden zelf maken uitbreiden (beknopte voorbeelden in template directory) deze diene in het bovenste directory te worden geplaats.
# accountnumbers.csv
# grootboeklijst.csv
# mapdecription-mappingtabel
# maptegenpartij-maptegenpartij.csv
# 
# 
# 

import pandas as pd
import os
import glob
import plotly.express as px
import numpy as np
from pathlib import Path


current_path = os.getcwd()


# ### Definitie Jaartal


# Deze variable stuurt voor welk jaartal er gekozen wordt.
jaar = 2023
jaar = str(jaar)


# De files zijn jaartal specifiek volgnummers en uitvoer
# De benaming is standaard als je download van de rabobank app

# Directory structuur 
OutputDir = "rabobank/" + jaar + "/"
DebugDir = "rabobank/" + jaar + "/"

# Invoer files
inVolgFile = "rabobank/" + jaar + "/volgnr_RABO" + jaar + ".csv"    #"rabobank/yyyy/volgnr_RABOxxxx.csv"
   # Mapping tabellen zijn universeel mapping tabellen en grootboek codes.
   # Prioriteit 1: volgnr_Rabo 2: Description 3: Tegenpartij
mappingdescription="rabobank/mapdescription-mappingtabel.csv"
maptegenpartij= "rabobank/maptegenpartij-maptegenpartij.csv"
grootboek="rabobank/grootboeklijst.csv"
accountnumbers = "rabobank/accountnumbers.csv"

#Uitvoer files
uitCSVFile = "rabobank/" + jaar + "/OUT_RABO" + jaar + ".csv"
Uitvoerfile ="rabobank/" + jaar + "/OUT_RABO" + jaar + "_final.xlsx"
UitvoerfilePartial  ="rabobank/" + jaar + "/OUT_RABO" + jaar + "_partial.xlsx"
uitDebugFile = "rabobank/" + jaar + "/nog_volgnummer_nodig_RABO" + jaar + ".csv"
UitTMP = "rabobank/" + jaar + "/temporary" + jaar + ".csv"

# Bericht kompleet
CompletedFile = OutputDir + "volgnr.completed"




# Overbodige files alvast weggooien
if os.path.exists(UitTMP):
    os.remove(UitTMP)

if os.path.exists(uitDebugFile):
    os.remove(uitDebugFile)

if os.path.exists(UitvoerfilePartial):
    os.remove(UitvoerfilePartial)


if os.path.exists(CompletedFile):
    os.remove(CompletedFile)    


# ### Inlezen en duplicaten weggooien
    
# Verkrijg het directory van de jupyter file.
path = current_path + "/rabobank/" + jaar + "/invoer"
os.chdir(path)

#Verzamel alle csv files die daarin worden gedumpt. Ze moeten beginnen met CSV_A voor de standaard rekeningen
all_filenames = [i for i in glob.glob("CSV_A*.csv")]
all_creditcardfiles = [i for i in glob.glob("CSV_CC*.csv")]


# Read in all RaboBank jaar csv Files from  subfolder rabobank the csv files stay untouched.
parts=[]
for f in all_filenames: 
    part=pd.read_csv(f, encoding = 'unicode_escape',
                        decimal="," ,parse_dates=["Datum","Rentedatum"])
    # Check number of columns
    # Getting shape of the df
    shape = part.shape
    # Printing Number of columns
    nKolom = shape[1]

    if nKolom == 26 :
        parts.append(part)
        

# Read in all creditcard csv Files from subfolder rabobank the csv files stay untouched.
cards=[]
for f in all_creditcardfiles: 
    card=pd.read_csv(f, encoding = 'unicode_escape',decimal = "," ,parse_dates= ["Datum"])
    # Check number of columns
    # Getting shape of the df
    shape = card.shape
    # Printing Number of columns
    nKolom = shape[1]


cards.append(card)
dfCC = pd.concat(parts)


# Pad weer terug zetten
path = current_path
os.chdir(path)


# Combine the Dataframes from each file into a single Dataframe
# pandas takes care of properly aligning the columns
dfIn = pd.concat(parts)

# ### Duplicaten verwijderen

dfIn=dfIn.drop_duplicates(subset=None, keep="first", inplace=False)

lengte_in = len(dfIn)

# ### Kolommen toevoegen en bewerken
#Toevoegen van maand en jaar
dfIn['year'] = pd.DatetimeIndex(dfIn['Datum']).year                                   
dfIn['month'] = pd.DatetimeIndex(dfIn['Datum']).month

#  Ik wil alleen de gegevens van jaar zien

dfIn= dfIn.loc[ (dfIn['year'] == int(jaar)) ]

# Kolommen weggooien
dfIn = dfIn.drop(['Omschrijving-2','Omschrijving-3','Reden retour','Oorspr bedrag',
                  'Oorspr munt','Batch ID'], axis=1)

#Kolommen hernoemen
dfIn.rename(columns={'IBAN/BBAN':'account','Omschrijving-1':'description',
                      'Tegenrekening IBAN/BBAN':'tegenrekening','Bedrag':'amount',
                     'Saldo na trn':'balance'
                    }, inplace = True)

### Filter creditcard afrekeningen eruit. Die worden separaat weer toegevoegd opgesplitst naar krediteur

#dfIn['description'] = dfIn['description'].astype('string')
#filter1='Kaartnummer: ****.****.****.9028   Zie rekeningoverzicht'
filter1='Zie rekeningoverzicht'
dfIn = dfIn.loc[~dfIn['description'].str.contains(filter1)]


# ## Processing CreditCard File


dfCC=dfCC.drop_duplicates(subset=None, keep="first", inplace=False)



dfCC = pd.concat(cards)
dfCC=dfCC.drop_duplicates(subset=None, keep="first", inplace=False)

# Kolommen weggooien
dfCC = dfCC.drop(['Creditcard Regel1','Creditcard Regel2','Oorspr bedrag','Oorspr munt','Oorspr bedrag','Koers'], axis=1)


#Kolommen hernoemen
dfCC.rename(columns={'Omschrijving':'description','Bedrag':'amount','Productnaam':'Naam tegenpartij',
                     'Tegenrekening IBAN':'account','Creditcard Nummer':'tegenrekening'}, inplace = True)

#Toevoegen van maand en jaar
dfCC['year'] = pd.DatetimeIndex(dfCC['Datum']).year                                   
dfCC['month'] = pd.DatetimeIndex(dfCC['Datum']).month

#Filter op het aangegeven jaar
dfCC= dfCC.loc[ (dfCC['year'] == int(jaar)) ]

### Filter "Verrekening vorig overzicht" in kolom description uit de file want dat voegt niks toe.
dfCC = dfCC.loc[dfCC['description'] != 'Verrekening vorig overzicht']

### Creditcard en bank overzichten samenvoegen
frames = [dfIn, dfCC]
dfIn = pd.concat(frames)

# ## Koppelen aan Grootboek/Kleinboek categorieen 
# ### Mapping tabellen laden

#Mapping table: sleutel(str) code(A1 etc)
dfMapDescription = pd.read_csv(mappingdescription)
# Mapping table: tegenpartij (str)  code(A1 etc)
dfMapTegenPartij = pd.read_csv(maptegenpartij)
# Er zitten wat floats tussen in de hoofdfile tussen die geef ik even een string waarde
dfIn["Naam tegenpartij"].fillna("leeg", inplace = True)


dfGrootboek = pd.read_csv(grootboek)
dfGrootboek.rename(columns={'Code':'category'}, inplace = True)
# #### Mapping description
map_code = pd.Series(dfMapDescription.code.values ,index=dfMapDescription.sleutel).to_dict()

# De magic function van stack overflow geeft een waarde of "--"

def extract_codes(row):
    # row is description
    for item in map_code:
        if item.lower() in row.lower():
            return map_code[item]
    return '--'

#Apply function on column description and add a column
dfIn['category'] = dfIn['description'].apply(extract_codes)


# ### Mapping tegenpartij
# #### functie kan zo worden hergebruikt voor de tegenpartij



map_code = pd.Series(dfMapTegenPartij.code.values ,index=dfMapTegenPartij.tegenpartij).to_dict()
#Apply function on column description and add a column
dfIn['category2'] = dfIn['Naam tegenpartij'].apply(extract_codes)
#dfIn.info()

# #### Er zijn nu wel twee category kolommen
# Let op dat de getallen 22 en 23 kolommen zijn nl categorie en categorie2 
# Daarbij is de mapping op description (categorie) leidend alleen als de mapping op tegenpartij iets oplevert wordt die waarde 
# toegekend aan de nieuwe kolom 'cat' 

def func(x):
    if x.iloc[23] != '--':
        return x.iloc[23]
    return x.iloc[22]


dfIn['cat'] = dfIn.apply(func,axis=1)
dfIn.pop('category')
dfIn.pop('category2')
# Rename the newly made category
dfIn.rename(columns={'cat':'category'}, inplace = True)


# ### Volg nummer file inlezen en koppelen aan de rabo csv
file_exists = os.path.exists(inVolgFile)

# Inlezen volgnummer file als de file niet betstaat maak een dataframe 
if file_exists:
    dfVolg=pd.read_csv(inVolgFile,encoding = 'unicode_escape')
    len(dfVolg)
    dfVolg=dfVolg.drop_duplicates(subset=['Volgnr'], keep="first", inplace=False)
    len(dfVolg)
else:
    # Create empty the pandas DataFrame
    dfVolg = pd.DataFrame(columns = ['Dummy', 'Volgnr', 'category'])

lengte_volgnr = len(dfVolg)


#dfVolg.info()

#Eerste is een overbodige kolom die gooi ik weg
dfVolg = dfVolg.drop(dfVolg.columns[[0]], axis=1)  


# ### Een merge uitvoeren om de volgnummer file te koppelen aan de hoofdfile 
dfIn=dfIn.merge(dfVolg, how="left", on="Volgnr")


# ### De volgnummer file is leidend als het goed is zijn er nu twee category kolommen aangemaakt
# category_x (22) is gebaseerd op de mapping tabellen die zijn niet perfect
# category_y (23) is gebaseerd op de volgnummer file die is goed maar het kan zijn dat er sommigen nog niet benoemd '--' of leeg zijn zijn.

#dfIn.to_csv(DebugDir+"0.csv")

def funcswitch(x):
    #if (pd.isnull(x[23])):
    if (pd.isnull(x.iloc[23])) | (x.iloc[23] == '--'):        
        return x.iloc[22]
    return x.iloc[23]


# Convert category_y en _X to string type to avoid errors in the funtion
dfIn['category_x'] = dfIn['category_x'].astype('string')
dfIn['category_y'] = dfIn['category_y'].astype('string')
#dfIn.info()

# The new column is called cat when it has a value from the mapping it will be  preceeded with "_"
# otherwise it gets the value from the volgnumber file
dfIn['cat'] = dfIn.apply(funcswitch,axis=1)

# #### En weer heb ik twee category kolommen
# delete the category columns they are superfluous now 
dfIn.pop('category_x')
dfIn.pop('category_y')
# Rename the newly made category
dfIn.rename(columns={'cat':'category'}, inplace = True)


# ### Geldstromen (Cashflows) toevoegen

# ### Maak een selectie van de geldstromen
# Please make sure you fill the CSV accountnumbers.csv 
# - Category 1 are the account which you pay with   (betaal rekeningen)
# - Category 2 are the account for savings money going out or in through category 1 (Spaarrekeningen)
# - Category 3 are the account for brokers  money going out or in through category 1 (Beleggingsrekeningen)
# 
# Category 1+2 are all bankaccounts (also the ones outside the rabobank)


# #### Maak een lijsten van alle rekeningen en geef er een specifieke naam aan¶

#Mapping table: sleutel(str) code(A1 etc)
# Lees de input file met rekening nummers
dfAccountnumbers = pd.read_csv(accountnumbers)

dfBetaalRekeningen = dfAccountnumbers.loc[dfAccountnumbers['category'] == 1 ]
BetaalRekeningen = dfBetaalRekeningen['accountno'].to_list()

dfSpaarRekeningen = dfAccountnumbers.loc[dfAccountnumbers['category'] == 2]
SpaarRekeningen = dfSpaarRekeningen['accountno'].to_list()

dfBeleggingsRekeningen = dfAccountnumbers.loc[dfAccountnumbers['category'] == 3 ]
BeleggingsRekeningen = dfBeleggingsRekeningen['accountno'].to_list()

Eigen_rekeningen = BetaalRekeningen + SpaarRekeningen
BeleggingsRekeningen


# #### Maak een list van de te benoemen cashflowcodes

#Cashflow codes let op de index begint bij 0 vandaar de dummy op nul
cashflow= ['','1-Inkomsten','2-Uitgaven',
           '3-Sparen', '4-Sparen opname',
           '5-Beleggen','6-Beleggen opname',
           "7-vestzak-broekzak"]


# Voeg te vullen kolom toe met een herkenbare default code
dfIn['cashflow']="-"

# ### 1 Hoofrekeningen (Beleggingen gaan alle vanaf hoofdrekening)¶

# #### 2-Uitgaven
# Alles wat afgeboekt wordt (<€0) van betaal rekeningen naar alles behalve eigen rekeningen zijn uitgaven
# Sparen en beleggen zit hier bij in maar wordt in een tweede slag afgevangen
dfIn.loc[(dfIn['account'].isin(Eigen_rekeningen)) &  
         (~dfIn['tegenrekening'].isin(Eigen_rekeningen)) &
         (dfIn['amount'] < 0),'cashflow']=cashflow[2]


# #### 1-Inkomsten

# Alles wat op de eigen rekeningen komt van alles behalve eigen rekeningen zijn inkomsten
# Sparen en beleggen zit hier bij in. Maar worden in stap 2 separaat afgevangen
dfIn.loc[(dfIn['account'].isin(Eigen_rekeningen)) & 
         (~dfIn['tegenrekening'].isin(Eigen_rekeningen)) &
         (dfIn['amount'] > 0),'cashflow']=cashflow[1]


# #### 7-Vestzak-broekzak

# Alles wat van en naar eigen rekeningen wordt overgeboekt is vestzak-broekzak 
# cq dubbel eigenlijk kan dit eruit
# Sparen en beleggen zit hier bij in. Maar worden in stap 2 separaat afgevangen
dfIn.loc[(dfIn['account'].isin(Eigen_rekeningen)) &  
         (dfIn['tegenrekening'].isin(Eigen_rekeningen)),
         'cashflow']=cashflow[7]


# ### 2. SpaarRekeningen; Sparen of Spaargeld Opnemen ?  

# #### 3-Sparen

#Afgeschreven (Bedrag<0) van Betaal Rekeningen gestort op spaar rekeningen noem ik 'Sparen'
#             (Bedrag>0) van spaar Rekeningen gestort op betaal rekeningen noem ik 'vestzak broekzak'
dfIn.loc[(dfIn['amount'] < 0) & (dfIn['account'].isin(BetaalRekeningen)) &
         (dfIn['tegenrekening'].isin(SpaarRekeningen)),'cashflow']=cashflow[3]


# #### 4-Sparen Opname

#Uitgekeerde (Bedrag>0) op Betaal Rekeningen afkomstig  van de spaar rekeningen noem ik 'Sparen Opnemen'
#            (Bedrag<0) van spaar Rekeningen  op betaal rekeningen noem ik 'vestzak broekzak'
dfIn.loc[(dfIn['amount'] > 0) & (dfIn['account'].isin(BetaalRekeningen))  & 
         (dfIn['tegenrekening'].isin(SpaarRekeningen)),'cashflow']=cashflow[4]


# ### 3. BeleggingsRekeningen: Beleggingen gaan alle (heen en terug) vanaf BetaalRekeningen 

# #### 5-Beleggen

# Bedragen die van de betaal rekeningen worden afgeschreven en naar de beleggings berekeningen 
# gaan zijn: 
# Beleggen
dfIn.loc[(dfIn['amount'] < 0) & (dfIn['account'].isin(BetaalRekeningen)) &
         (dfIn['tegenrekening'].isin(BeleggingsRekeningen)),'cashflow']=cashflow[5]


# #### 6-Beleggen Opname

# Bedragen die op de betaalrekeningen worden bijgeschreven en afkomstig zijn van de beleggings berekeningen 
# gaan zijn beleggingsgeld wat wordt opgenomen naar actieve rekening
# Beleggen Opname
dfIn.loc[(dfIn['amount'] > 0) & (dfIn['account'].isin(BetaalRekeningen)) &
         (dfIn['tegenrekening'].isin(BeleggingsRekeningen)),'cashflow']=cashflow[6]


# ### De grootboek codes koppelen aan de category codes

dfGrootboek = pd.read_csv(grootboek)
dfGrootboek.rename(columns={'Code':'category'}, inplace = True)

dfIn=dfIn.merge(dfGrootboek, how="left", on="category")


# #### Benodigde Pivot tables maken van de geldstromen/grootboek/kleinboek
# 
# We houden de interne kasstromen (vestzak-broekzak) erbuiten

dfPivot=dfIn.loc[(~dfIn['cashflow'].str.contains(cashflow[7]))]
#dfPivot.to_csv(DebugDir + "dfPivot.csv")
#dfPivot.head()

# Cashflow op maand basis
pivot_table1 = pd.pivot_table(dfPivot, values='amount', index=['cashflow'],columns=['month'], aggfunc='sum',
                      fill_value=0, margins=True)
                      
#Cashflow op jaar basis op jaarbasis exclusief vestzak broekzak (niet erg interessant)
pivot_table2 = pd.pivot_table(dfPivot, values='amount', index=['cashflow'],columns=['year'], aggfunc='sum',
                      fill_value=0, margins=False)

# Grootboek op jaarbasis exclusief vestzak broekzak
pivot_table3 = pd.pivot_table(dfPivot, values='amount', index=['Grootboek','Kleinboek'],columns=['month'], aggfunc='sum',
                      fill_value=0, margins=True)


# ### Files ingelezen en grootboek categorieen toegevoegd 

# ### Debug output
# Alles wat nog geen grootboek code heeft gehad wordt even apart in een file gezet 

dfUitDebug=dfIn.loc[(dfIn['category'].str.contains('--'))]
dfUitDebug = dfUitDebug[['account','Datum','balance','year','month',
'tegenrekening','Naam tegenpartij',
'description','amount','Volgnr','category','cashflow']]

completed = False
if len(dfUitDebug) >0 :
    dfUitDebug.to_csv(uitDebugFile)
else:
    # Write empty dataframe and give a message in the filename
    completed = True
    dfUitDebug.to_csv(OutputDir + "volgnr.completed")


# ## Uitvoer

dfIn=dfIn[['account','Munt','BIC','Volgnr','Datum','Rentedatum','amount','balance',
'tegenrekening','Naam tegenpartij','Naam uiteindelijke partij','Naam initiërende partij',
'BIC tegenpartij','Code','Transactiereferentie','Machtigingskenmerk','Incassant ID','Betalingskenmerk',
'description','Koers','year','month','category','Grootboek','Kleinboek','cashflow']]

# Create a Pandas Excel writer using XlsxWriter as the engine.
if completed:
    writer = pd.ExcelWriter(Uitvoerfile, engine='xlsxwriter',datetime_format='yyyy mm dd', date_format='yyyy mm dd')
else:
    writer = pd.ExcelWriter(UitvoerfilePartial, engine='xlsxwriter',datetime_format='yyyy mm dd', date_format='yyyy mm dd')

#dfIn=dfIn.style.set_properties(**{'text-align': 'left'})

# Wegschrijven naar Excel ongeformatteerd
dfIn.to_excel(writer, sheet_name='0-master', index = False )

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['0-master']

# Get the dimensions of the dataframe
(max_row, max_col) = dfIn.shape

# Create a list of column headers, to use in add_table()
column_settings = [{'header': column} for column in dfIn.columns]

# Add the Excel table structure. Pandas will add the data.
worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

# Add some cell formats.
# Datum jjjj-mm-dd;@
# Currency € #,##0.00
# Getal 0

formatCur = workbook.add_format({'num_format': '€ #,##0.00'})
formatDate = workbook.add_format({'num_format': 'jjjj-mm-dd;@'})
formatBold =  workbook.add_format()
#formatBold =  workbook.add_format({'bold': True})
formatBold.set_align('left')

#formatInt = workbook.add_format({'num_format': '0'})

# Zet de format op de kolommen
# worksheet.set_column('G:XFD', None, None, {'hidden': True})
worksheet.set_column(0, 0, 19)
worksheet.set_column(2, 2, 9)
worksheet.set_column(4, 5, 12)
worksheet.set_column(6, 7, 12, formatCur)
worksheet.set_column(8, 8, 23)
worksheet.set_column(9, 9, 50)
worksheet.set_column(10, 10, 20, None, {'hidden': True})
worksheet.set_column(11, 11, 32, None, {'hidden': True})
worksheet.set_column(12, 12, 12)
worksheet.set_column(13, 13, 8)
worksheet.set_column(14, 14, 37, None, {'hidden': True})
worksheet.set_column(15, 15, 36, None, {'hidden': True})
worksheet.set_column(16, 16, 30, None, {'hidden': True})
worksheet.set_column(17, 17, 15, None, {'hidden': True})
worksheet.set_column(18, 18, 110)
worksheet.set_column(19, 19, 15, None, {'hidden': True})
worksheet.set_column(23, 23, 14)
worksheet.set_column(24, 24, 27)
worksheet.set_column(25, 25, 15)


# ### Ad format to the pivot and add them to the excel file as seperate sheets

# Wegschrijven van de pivots naar dezelfde Excelfile 
pivot_table1.to_excel(writer, sheet_name='1-Pivot', index = True )
#pivot_table2.to_excel(writer, sheet_name='2-Pivot', index = True )
pivot_table3.to_excel(writer, sheet_name='3-Pivot', index = True )

worksheet = writer.sheets['3-Pivot']
worksheet.set_column(0, 0, 20,formatBold)   # Bold en left Werkt niet op de index kolommen
worksheet.set_column(1, 1, 30,formatBold)   # Bold en left Werkt niet op de index kolommen
for col in range(2,15):
    worksheet.set_column(col, col, 13, formatCur)

worksheet = writer.sheets['1-Pivot']
worksheet.set_column(0, 0, 20,formatBold)   # Bold en left Werkt niet op de index kolommen
for col in range(1,14):
    worksheet.set_column(col, col, 13, formatCur)


# ### Sheets for saldo verloop

for rekening in Eigen_rekeningen:
    #worksheet = writer.sheets[rekening]
    dfUit = dfIn.loc[(dfIn['account'].str.contains(rekening)), ['account','Datum','amount','balance']]
    dfUit.to_excel(writer, sheet_name=rekening, index = True )
    worksheet = writer.sheets[rekening]
    worksheet.set_column(1, 1, 20)
    worksheet.set_column(2, 2, 15)
    worksheet.set_column(3, 3, 13, formatCur)
    worksheet.set_column(4, 4, 13, formatCur)
  

# Close the Pandas Excel writer and output the Excel file.
writer.close()
