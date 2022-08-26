# Imports
import json
import xlwt

# Initialiseer xlwt en json
f = open('dealers_faber.json')
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Dealers")
data = json.load(f)

# Indexes
idIndex = 0
naamIndex = 1
landIndex = 2
adresIndex = 3
websiteIndex = 4
telefoonIndex = 5
afspraakIndex = 6

# Headers
sheet1.write(0, idIndex, 'ID')
sheet1.write(0, naamIndex, 'Naam')
sheet1.write(0, adresIndex, 'Adres')
sheet1.write(0, landIndex, 'Land')
sheet1.write(0, websiteIndex, 'Website')
sheet1.write(0, telefoonIndex, 'Telefoon')
sheet1.write(0, afspraakIndex, 'Alleen op afspraak')

# Index
row = 1

# Voor elke value
for key, value in data.items():
    # Land
    land = str(value['address']).split('\n')[-1]
    sheet1.write(row, landIndex, land)

    # Adres
    adres = str(value['address'])
    if adres != '':
        # haal de Naam uit het adres
        adres = adres.replace(value['title'], '')
        # replace alle Enters met kommas
        adres = adres.replace('\n', ', ')
        # replace eerste komma met niks
        adres = adres.replace(', ', '', 1)
        # haal land uit adres
        landMetKomma = ', ' + land
        adres = adres.replace(landMetKomma, '', 1)
        if adres == '':
            adres = 'Geen gegevens'
    elif adres == '':
        adres = 'Geen gegevens'

    sheet1.write(row, adresIndex, adres)

    # Op afspraak
    isAlleenOpAfspraak = ''
    if value['appointment_only'] == True:
        isAlleenOpAfspraak = 'Ja'
    elif value['appointment_only'] == False:
        isAlleenOpAfspraak = 'Nee'
    else:
        isAlleenOpAfspraak = 'Geen gegevens'
    sheet1.write(row, afspraakIndex, isAlleenOpAfspraak)

    # Overige gegevens
    sheet1.write(row, idIndex, value['counter'])
    sheet1.write(row, naamIndex, value['title'])

    website = value['website']
    if value['website'] == None:
        website = 'Geen gegevens'

    sheet1.write(row, websiteIndex, website)

    phone = str(value['phone'])
    if value['phone'] == '':
        phone = 'Geen gegevens'
    sheet1.write(row, telefoonIndex, value['phone'])

    row = row + 1

# Sluit bestand af
f.close()
# Sla excel bestand op en geef het een naam
book.save('dealers_faber.xls')
