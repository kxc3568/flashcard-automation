import xlrd
import re
from splinter import Browser
import urllib.request

filename = "~/Downloads/ycm.xlsx"
wb = xlrd.open_workbook(filename)
sheet = wb.sheet_by_index(0)
br = Browser()
br.visit("https://www.cardmaker.net/yugioh/")
form_ids = [
    "name", "cardtype", "subtype", "attribute", "level", "trapmagictype", "rarity", "picture", "circulation", "set1",
    "set2", "type", "carddescription", "atk", "def", "creator", "year", "serial"
]
opener = urllib.request.build_opener()
opener.addheaders = [("User-Agent", "Mozilla/5.0")]
urllib.request.install_opener(opener)

for r in range(2, sheet.nrows):
    for c in range(1, sheet.ncols):
        form_id = form_ids[c-1]
        form_value = sheet.cell_value(r, c)
        if not form_value:
            break
        if c == 3:
            form_value = form_value.lower()
        elif c == 5 or c == 14 or c == 15 or c == 17:
            form_value = str(int(form_value))
        try:
            br.find_by_id(form_id).first.fill(form_value)
        except Exception:
            br.find_by_id(form_id).first.select(form_value)
    br.find_by_id("generate").first.click()
    if sheet.cell_value(r, 1):
        src = br.find_by_id("card").first["src"]
        urllib.request.urlretrieve(src, sheet.cell_value(r, 1) + ".jpeg")