# Imports
import json
import xlwt
import pandas as pd
import re

# Init xlwt
book = xlwt.Workbook(encoding="utf-8")
sheet = book.add_sheet("Blad1")

# Script

def split_on_letter(s):
    match = re.compile("[^\W\d]").search(s)
    if(match != None):
        return [s[:match.start()], s[match.start():]]
    else:
        return [s, '']

exact_df = pd.read_excel(
    'C:\Files\Code\GitHub\Thermocet\De Ridder Import\input\exact_export.xls', sheet_name='Sheet1')
afladr_df = pd.read_excel(
    'C:\Files\Code\GitHub\Thermocet\De Ridder Import\import_afladr.xlsx', sheet_name='Blad1')
nr_file = pd.read_excel(
    'C:\Files\Code\GitHub\Thermocet\De Ridder Import\\nr_file.xls', sheet_name='Sheet')

nr_file = nr_file[nr_file['Debiteurennummer'].notnull()]

baseIndex = 4
start_index = 0
matches = 0
no_matches = []

# for each row in import file
for index1, row1 in afladr_df.iterrows():

    if(index1 > 3):
        # print('index:', index1)
        # print('start index:', start_index)
        import_naam = row1[1]
        # print('huidige:', import_naam, 'index:', index1)

        exact_name_index = -1

        # for each row in exact file
        for index2, row2 in exact_df.iterrows():

            if(index2 < start_index):
                continue

            # if 'Debiteur' is in naam, remove it
            if(not pd.isnull(row2['Stamgegevens Debiteuren'])):
                if('Debiteur:' in row2['Stamgegevens Debiteuren']):

                    deb_nr = row2['Stamgegevens Debiteuren'].split(' ')[1]
                    deb_naam = ' '.join(
                        row2['Stamgegevens Debiteuren'].split(' ')[2:])

                    # zoek naar debiteurnummer in debiteuren file
                    for index3, row3 in nr_file.iterrows():

                        if(not pd.isnull(row3['Debiteurennummer'])):
                            # als match, pak naam
                            if(deb_nr == row3['Debiteurennummer']):

                                nr_file_naam = row3['Naam']

                                if(import_naam == nr_file_naam):

                                    matches += 1
                                    exact_name_index = index2

                                    break
                    else:
                        continue
                    break

        if(exact_name_index != -1):
            # print('match gevonden voor', import_naam,
            #       'op index', exact_name_index)


            start_index = exact_name_index + 15

            # write to excel
            current_deb_data = exact_df[exact_name_index-3:exact_name_index+10]
            
            # get if multiple afladr
            # print(current_deb_data['Stamgegevens Debiteuren'])
            mylist = list(current_deb_data['Stamgegevens Debiteuren'])
            amt_of_afladr = 0
            for i in range(len(mylist)):
                if(mylist[i] == 'Afleveradres'):
                    amt_of_afladr += 1
            
            if(amt_of_afladr > 1):
                print('meerdere afladr voor', mylist[3].replace('Debiteur: ', ''), 'op index', exact_name_index, 'aantal:', amt_of_afladr)
                continue

            afl_adres_land = current_deb_data['Unnamed: 24'].iloc[5].split(' ')[0]

            aflever_adres = current_deb_data['Gebruiker'].iloc[5]
            if(aflever_adres != None and aflever_adres != ' ' and not pd.isnull(aflever_adres)):
                aflever_adres = re.split(r'(^[^\d]+)', aflever_adres)[1:]
                if(aflever_adres != []):
                    afl_adres_naam = aflever_adres[0].strip()
                    afl_adres_nr_en_toev = aflever_adres[1].strip()

                    afl_adres_toev = split_on_letter(afl_adres_nr_en_toev)[1].strip()
                    afl_adres_nr = split_on_letter(afl_adres_nr_en_toev)[0].strip()
                    
                    afl_adres_postcode_en_plaatsnaam = current_deb_data['Unnamed: 15'].iloc[5]
                    afl_adres_postcode = afl_adres_postcode_en_plaatsnaam.split(' ')[
                        0:2]
                    afl_adres_postcode = ' '.join(afl_adres_postcode).strip()
                    afl_adres_plaatsnaam = ' '.join(
                        afl_adres_postcode_en_plaatsnaam.split(' ')[2:]).strip()

                    # print('aflever adres:', afl_adres_naam, afl_adres_nr,
                    #     afl_adres_toev, afl_adres_postcode, afl_adres_plaatsnaam, afl_adres_land)

                    # write to excel
                    sheet.write(baseIndex, 0, import_naam)
                    sheet.write(baseIndex, 1, afl_adres_naam)
                    sheet.write(baseIndex, 2, afl_adres_nr)
                    sheet.write(baseIndex, 3, afl_adres_toev)
                    sheet.write(baseIndex, 4, afl_adres_postcode)
                    sheet.write(baseIndex, 5, afl_adres_plaatsnaam)
                    sheet.write(baseIndex, 6, afl_adres_land)

            baseIndex += 1


        else:
            print('geen match gevonden voor', import_naam)
            no_matches.append(import_naam)


print('TOTAAL: ' + str(matches) + ' matches')

print('GEEN MATCHES:')
for i in no_matches:
    print('-', i)


book.save("output.xls")

# 147
