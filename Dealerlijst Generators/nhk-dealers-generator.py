# Imports
import json
import xlwt

# Initaliseer xlwt en json
f = open('input/dealers_nhk.json')
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Dealers")
data = json.load(f)

# Indexes
idIndex = 0
naamIndex = 1
adresIndex = 2
postcodeIndex = 3
plaatsnaamIndex = 4
provincieIndex = 5
landIndex = 6
websiteIndex = 7
emailIndex = 8
telefoonIndex = 9

# Headers
sheet1.write(0, idIndex, '#')
sheet1.write(0, naamIndex, 'Naam')
sheet1.write(0, adresIndex, 'Adres')
sheet1.write(0, plaatsnaamIndex, 'Plaatsnaam')
sheet1.write(0, provincieIndex, 'Provincie')
sheet1.write(0, postcodeIndex, 'Postcode')
sheet1.write(0, landIndex, 'Land')
sheet1.write(0, websiteIndex, 'Website')
sheet1.write(0, emailIndex, 'E-mail')
sheet1.write(0, telefoonIndex, 'Telefoon')

# Index
row = 1

# Dealers uit data object
dealers = data

# Voor elke dealer
for dealer in dealers:
    # Alle gegevens
    sheet1.write(row, idIndex, row)
    sheet1.write(row, naamIndex, dealer['store'])
    sheet1.write(row, landIndex, dealer['country'])
    sheet1.write(row, adresIndex,
                 dealer['address'])
    sheet1.write(row, postcodeIndex,
                 dealer['zip'])
    sheet1.write(row, plaatsnaamIndex, dealer['city'])
    sheet1.write(row, provincieIndex, dealer['state'])
    sheet1.write(row, websiteIndex, dealer['url'])
    sheet1.write(row, emailIndex, dealer['email'])
    sheet1.write(row, telefoonIndex, dealer['phone'])

    row = row + 1

# Sluit bestand
f.close()
# Sla excel bestand op en geef het een naam
book.save('dealers_nhk.xls')
print('Completed')
