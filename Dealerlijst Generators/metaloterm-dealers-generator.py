# Imports
import json
import xlwt

# Initialiseer xlwt en json
f = open('input/dealers_metaloterm_wereldwijd.json')
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Dealers")
data = json.load(f)

# Indexes
idIndex = 0
naamIndex = 1
adresIndex = 2
landIndex = 3
postcodeIndex = 4
plaatsnaamIndex = 5
websiteIndex = 6
telefoonIndex = 7
latIndex = 8
longIndex = 9

# Headers
sheet1.write(0, idIndex, '#')
sheet1.write(0, naamIndex, 'Naam')
sheet1.write(0, adresIndex, 'Adres')
sheet1.write(0, landIndex, 'Land')
sheet1.write(0, postcodeIndex, 'Postcode')
sheet1.write(0, websiteIndex, 'Website')
sheet1.write(0, telefoonIndex, 'Telefoon')
sheet1.write(0, latIndex, 'Breedtegraad')
sheet1.write(0, longIndex, 'Lengtegraad')

row = 1
for dealer in data:
    # Alle gegevens
    sheet1.write(row, idIndex, row)
    sheet1.write(row, naamIndex, dealer['company'])
    sheet1.write(row, adresIndex, dealer['address'])
    sheet1.write(row, landIndex, dealer['country'])
    sheet1.write(row, postcodeIndex, dealer['zipcode'])
    sheet1.write(row, plaatsnaamIndex, dealer['city'])
    sheet1.write(row, websiteIndex, dealer['website'])
    sheet1.write(row, telefoonIndex, dealer['phone'])
    sheet1.write(row, latIndex, dealer['latlong']['latitude'])
    sheet1.write(row, longIndex, dealer['latlong']['longitude'])

    row = row + 1

# Sluit bestand
f.close()
# Sla excel bestand op en geef het een naam
book.save('dealers_metaloterm_wereldwijd.xls')
print('Completed')
