import time
import requests
from lxml import html
import xlsxwriter

begin = time.time()

## Modify this part as needed
#  ================

signs = ['INTC']

dates = [
    '1426809600', '1429228800', '1431648000', '1434672000', '1437091200',
    '1440115200']
    
excel = xlsxwriter.Workbook('stock.xlsx') # put filename here

left_col = "//div[@id='optionsCallsTable']//tbody/tr"

path = "//div[@id='optionsCallsTable']//tbody/tr[{}]/td/*//text()"

site = 'https://finance.yahoo.com/q/op?s={}&date={}' # Remember to call .format
                                                     # if using somewhere else
## Don't change this
#  =================

ref = 0

sheet = excel.add_worksheet()

for sign in signs:
    for date in dates:
        print 'Getting', sign, str(date) + '...',
        page = requests.get(site.format(sign, date))
        tree = html.fromstring(page.text)
        left_data = tree.xpath(left_col)
        table = []
        for row_n in range(len(left_data)):
            table.append(tree.xpath(path.format(row_n + 1)))
        for row_n, row in enumerate(table):
            for col_n, cell in enumerate(row):
                sheet.write(row_n + ref, col_n, cell)
        ref += len(table) + 1   # Counting starts at 0
        print 'Done'
        
end = time.time()

print 'Completed in', round(end - begin, 4), 'seconds'
