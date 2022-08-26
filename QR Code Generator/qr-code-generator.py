# URLs
# https://www.xml-sitemaps.com/
# https://png2pdf.com/nl/

# Geef de bestandsnaam name-TAAL.png als naam

import qrcode
with open('urllist-it.txt', 'r') as f:
    urls = f.readlines()
    for i in urls:
        i = i.strip()
    for i in urls:
        img = qrcode.make(i)
        type(img)
        name = i.split('/')[-2]
        img.save(name + '-it.png')