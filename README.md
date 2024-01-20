# rabobank_csv
Reads the RaboBank csv file and translates it to an Excel workbook.
I think this is only interesting for Dutch people who actually use the RaboBank so I will continue in Dutch.

Ieder jaar download ik een csv file van de rabobank. Ik werk met een MacBook en die kan de csv files aardig goed inlezen. Maar als je wat beter wilt kijken naar al die rekeningen en transacties ontkom je niet aan Excel.
Tot 2021 deed ik dit allemaal handmatig met als gevolg dat de handleiding die ik schreef iedere keer weer veranderde omdat ik het eigenlijk niet meer wist hoe ik het de vorige keer had gedaan.
Samengevat wil ik de volgende features hebben in een uiteindelijke Excel File.

- Datums zijn datums, valuta zijn valuta 
- Alle transacties van al mijn bankrekening in een tabel zodat ik gauw kan filteren.
- Alle transacties worden gecategoriseerd naar Inkomsten, Uitgaven, Sparen en Beleggen
- 
Met enig handmatig aanmaken van files wil ik het volgende krijgen:
- Alle transacties bevatten/krijgen een categorie die ik grootboek en kleinboek heb genoemd.
- Additionele worksheet per rekening om een saldo verloop te genereren
- Additionele worksheet met een pivot tabel boekingen per grootboek/kleinboek per maand.
- Mogelijkheid om meerdere csv files op te nemen zodat je op maand of weekbasis kunt werken.
- Mogelijke duplicaten eruit fileteren

--------------------------------------------------------------------------------------------
Directory structuur **bold** zijn directories
- **rabobank**
	- RaboBank_master 20xx.ipynb  De jupyter note book file of py file
	- **rabobank**
		- Grootboeklijst.csv			Bevat de grootboek codes bv: A1 Inkomsten Loon 
		- mapdescription-mappingtabel.csv	Bevat stukjes tekst in kolom description die verbonden zijn aan een grootboek code
		- maptegenpartij-maptegenpartij.csv	Bevat stukjes tekst in kolom tegenpartij die verbonden zijn aan een grootboek code
                - accountnumbers.csv 			Bevat de rekening nummers 1: betaalrekeningen 2: Spaarrekeningen 3: beleggingen/lijrentes
		- **2024**
			- volgnr_RABO2024.csv		Handmatig koppel een volgnummer aan een kleinboek code. 
			- **invoer**
				- _all RaboBank csv_a files which contain transaction. from that year
				- for example: CSV_A_20220205_124700.csv
    				
				- _all RaboBank csv_cc files which contain creditcard transactions
				- for example: CSV_CC_20220205_124700.csv			- 
		- **2025** etc



Wat je als gebruiker wel moet doen is het volgende:

Eenmalige actie:

- Aangeven welke jouw eigen rekeningen zijn en wat voor rekeningen het zijn (file: accountnumbers.csv)
- Welke daarvan zijn spaarekeningen en welke betaal rekeningen
- Als je beleggingsrekeningen hebt geef die dan ook op
- Je directory structuur maken of die van mij overnemen
- Creeer mapping tabellen:

	- mapdescription-mappingtabel.csv		Een grootboek code voor iedere description (iets wat in description voorkomt)
	- maptegenpartij-maptegenpartij.csv		Een grootboek code voor iedere tegenpartij (iets wat in tegenpartij voorkomt)
  	- volgnr_RABO2022.csv	(optioneel)		Een kleinboek code voor ieder Volgnummer 
		
Structuur CSV Grootboek lijst (kun je naar eigen behoefte invullen, basis file is toegevoegd.)

<img width="296" alt="image" src="https://user-images.githubusercontent.com/73278171/153902978-2462cab3-9441-4e9d-af9e-5c091b932fb3.png">

Structuur CSV mapdescription-mappingtabel.csv  (moet je naar eigen behoefte invullen, basisfile is toegevoegd )

<img width="234" alt="image" src="https://user-images.githubusercontent.com/73278171/153903577-e9de77ff-e2ae-4718-a1e5-dbfbf3e5f7c4.png">

Structuur CSV maptegenpartij-maptegenpartij.csv  (moet je naar eigen behoefte invullen, basis file is toegevoegd)

<img width="151" alt="image" src="https://user-images.githubusercontent.com/73278171/153903775-d06a7a8b-1625-40ce-a513-bce2f4974767.png">


Structuur CSV volgnr_RABO2022.csv  (moet je naar eigen behoefte invullen, basis file is toegevoegd)	

<img width="120" alt="image" src="https://user-images.githubusercontent.com/73278171/153904320-86b2d5d5-5b8f-4b84-8c28-ddf984617eee.png">
	
	
Het idee is dat dat de veelvoorkomende transacties worden afgevangen (grootboek code aangehangen) door de beide mapping tabellen. Waarbij decription-table prioriteit heeft boven tegenpartij-tabel.
Gaandeweg door toevoegingen van de gebruiker worden deze mapping tabellen steeds beter, maar er zullen altijd weer transacties zijn die er iets buiten vallen.
Daar komt het volgnummer bij kijken deze file heeft de meeste prioriteit deze is gaandeweg makkelijk te genereren uit de bestaand tabellen.
Ieder run wordt er een file aangmaakt die heet _nog_volgnummer_nodig_RABO2022.numbers_ die je kunt gebruiken al eerst volgnummer file door voor ontbrekende volgnummers de grootboek code toe te voegen.

Al met al zit er nog wat handwerk aan maar als je het net als ik op maand basis bijhoudt valt het erg mee.
De gegenereerde excel file is erg handig voor verdere analyse 

Succes!



