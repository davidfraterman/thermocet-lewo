# Imports
import json
import xlwt

# Initaliseer xlwt en json
f = open('dealers_bellfires.json')
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Dealers")
data = json.load(f)

# Indexes
idIndex = 0
naamIndex = 1
landIndex = 2
adresIndex = 3
postcodeIndex = 4
plaatsnaamIndex = 5
websiteIndex = 6
emailIndex = 7
telefoonIndex = 8

# Headers
sheet1.write(0, idIndex, 'ID')
sheet1.write(0, naamIndex, 'Naam')
sheet1.write(0, adresIndex, 'Adres')
sheet1.write(0, postcodeIndex, 'Postcode')
sheet1.write(0, plaatsnaamIndex, 'Plaatsnaam')
sheet1.write(0, landIndex, 'Land')
sheet1.write(0, websiteIndex, 'Website')
sheet1.write(0, emailIndex, 'E-mail')
sheet1.write(0, telefoonIndex, 'Telefoon')

# Index
row = 1

# Dealers uit data object
dealers = data["dealers"]

# Voor elke dealer
for dealer in dealers:
    # Alle gegevens
    sheet1.write(row, idIndex, dealer['id'])
    sheet1.write(row, naamIndex, dealer['title'])
    sheet1.write(row, landIndex, dealer['contactInformation']['country'])
    sheet1.write(row, adresIndex, dealer['contactInformation']['streetAddress'])
    sheet1.write(row, postcodeIndex, dealer['contactInformation']['postalCode'])
    sheet1.write(row, plaatsnaamIndex, dealer['contactInformation']['city'])
    sheet1.write(row, websiteIndex, dealer['contactInformation']['website'])
    sheet1.write(row, emailIndex, dealer['contactInformation']['email'])
    sheet1.write(row, telefoonIndex, dealer['contactInformation']['phone'])

    row = row + 1

# Sluit bestand
f.close()
# Sla excel bestand op en geef het een naam
book.save('dealers_bellfires.xls')
