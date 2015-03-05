from lxml import html
import requests
import xlsxwriter as xlsx
import time

# TO BE ADDED
# [X] Time tracked during download (double alliteration FTW)
# [ ] Extra columns (look at example spreadsheet)

# --------------------------------------------------------------------------- #
signs = ['INTC', 'AAPL', 'MSFT', 'GOOG']

dates = ["1426809600", "1429228800", "1431648000", "1434672000", "1437091200",
         "1440115200"]
# --------------------------------------------------------------------------- #

headers = ['Sign', 'Date', 'Strike', 'Contract Name', 'Last', 'Bid', 'Ask',
           'Change', '%Change', 'Volume', 'Open Interest',
           'Implied Volatility']

formatting = ['str', 'int', 'float', 'str', 'float', 'float', 'float', 'float',
              'percent', 'int', 'int', 'percent']


def getStock(sign, date):
    """Returns Yahoo Finance future info."""

    # Download stock data into HTML tree
    print("Downloading data for {} at {}...".format(sign, date))
    page = requests.get(
        "https://finance.yahoo.com/q/op?s={}+Options".format(sign))
    tree = html.fromstring(page.text)

    # Create headers
    headers = ['Strike', 'Contract Name', 'Last', 'Bid', 'Ask', 'Change',
               '%Change', 'Volume', 'Open Interest', 'Implied Volatility']

    # Get all text in each row
    data = []
    rows = tree.xpath('//div[@id="optionsCallsTable"]//tbody/tr')
    for row in range(len(rows)):
        toAppend = [sign, date]

        toAppend.extend(tree.xpath(
            '//div[@id="optionsCallsTable"]//tbody/tr[{}]/td/*//text()'.format(
                row + 1)))

        data.append(toAppend)

    # Finish and return
    if data is not []:
        return data


def getStocks(signs, dates):
    """Returns data for multiple stocks in list form."""

    data = []
    for sign in signs:
        for date in dates:
            data.extend(getStock(sign, date))

    return data


def printToExcel(data, filepath, x_offset=0, y_offset=0):
    """Print a 2-dimensional list to an Excel file."""

    workbook = xlsx.Workbook(filepath)
    worksheet = workbook.add_worksheet()
    percent = workbook.add_format({'num_format': '\%0.00'})
    float = workbook.add_format({'num_format': '##0.00'})

    print "Writing data..."
    for row, row_data in enumerate(data):
        for column, cell in enumerate(row_data):
            if formatting[column] is 'percent':
                worksheet.write(row + y_offset, column + x_offset, cell,
                                percent)

            elif formatting[column] is 'float':
                worksheet.write(row + y_offset, column + x_offset, cell, float)

            else:
                worksheet.write(row + y_offset, column + x_offset, cell)

    workbook.close()

# --------------------------------------------------------------------------- #
start = time.time()
print 'Starting at {}'.format(start)
data = getStocks(signs, dates)
download_end = time.time()
print 'Download completed in {} seconds'.format(round(download_end - start, 2))

for row in data:
    for i in range(len(row)):
        if formatting[i] is 'percent':
            row[i] = float(row[i][:-1])
        else:
            row[i] = eval('{}(row[i])'.format(formatting[i]))


data.insert(0, headers)

printToExcel(data, 'stockdata.xlsx')

print 'Completed in {} seconds'.format(round(time.time() - start, 2))
