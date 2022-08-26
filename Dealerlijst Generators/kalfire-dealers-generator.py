# Imports
try:
    from BeautifulSoup import BeautifulSoup
except ImportError:
    from bs4 import BeautifulSoup

import xlwt

# Init xlwt
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Dealers")

# Indexes
idIndex = 0
naamIndex = 1
adresIndex = 2
plaatsnaamIndex = 3
websiteIndex = 4
telefoonIndex = 5

# Headers
sheet1.write(0, idIndex, 'ID')
sheet1.write(0, naamIndex, 'Naam')
sheet1.write(0, adresIndex, 'Adres')
sheet1.write(0, plaatsnaamIndex, 'Plaatsnaam')
sheet1.write(0, websiteIndex, 'Website')
sheet1.write(0, telefoonIndex, 'Telefoon')

# Open html bestand en voed het aan beautiful soup
with open("kalfire.html", encoding="utf8") as fp:
    soup = BeautifulSoup(fp)

# Counter voor row en ID waarde
counter = 1

# Voor elk dealer <div> component met class dealers-dealer
for dealer in soup.find_all('div', {'class': 'dealers-dealer'}):
    # Zoek naar naam in <a> tag, neem tekst
    naam = dealer.find('a').text
    naam = naam.lstrip()
    # Zoek naar adres in <address>, neem tekst
    locatie = dealer.find('address').text
    # Pak plaats en adres van locatie array
    plaats = locatie.split('\n')[2]
    plaats = plaats.lstrip()
    adres = locatie.split('\n')[1]
    adres = adres.lstrip()
    # Zoek website in <a> tag met rel='', pak de href waarde
    website = dealer.find('a', {'rel': 'external nofollow'})['href']
    website = website.lstrip()
    if website == 'http://':
        website = 'Geen website'
    # Pak telefoon nummer in <div> tag met dealer-details class, pak de tekst
    # Neem de derde + haal T: weg
    tel = dealer.find('div', {'class': 'dealer-details'}).text
    tel = tel.split('\n')[2]
    tel = tel.replace("T: ", "")
    tel = tel.lstrip()
    if tel == "":
        tel = 'Geen telefoonnummer'

    sheet1.write(counter, idIndex, counter)
    sheet1.write(counter, naamIndex, naam)
    sheet1.write(counter, adresIndex, adres)
    sheet1.write(counter, plaatsnaamIndex, plaats)
    sheet1.write(counter, websiteIndex, website)
    sheet1.write(counter, telefoonIndex, tel)

    counter = counter + 1

# Sla bestand op en geef het een naam
book.save('dealers_kalfire.xls')