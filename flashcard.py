import xlrd
import re
import mechanize

filename = "~/Downloads/ycm.xlsx"
wb = xlrd.open_workbook(filename)
sheet = wb.sheet_by_index(0)

br = mechanize.Browser()
br.set_handle_robots(False)
br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]
br.open("https://www.cardmaker.net/yugioh/")
br.select_form(nr=2)

form_labels = [sheet.cell_value(1, i) for i in range(1, sheet.ncols)]
for i in range(2, sheet.nrows):
    for j in range(1, sheet.ncols):
        pass