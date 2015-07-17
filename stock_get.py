## To-Do List

# [X] Separate sheets for sectors
# [X] Add easier way to input dates
# [X] Add separate files for dates and signs
# [ ] Automated date validity check

## Import Modules


from datetime import datetime
import time
import os
from os.path import expanduser
import socket
from subprocess import call
import sys

try:
    from lxml import html
    import openpyxl
    import requests
    from textwrap import fill
    import xlsxwriter
except ImportError as e:
    e = str(e)
    print(e)
    module = e[e.find("'") + 1:e.find("'", -1)]
    if 'y' in input('Module "{0}" is not installed. Would you like to install it? (y/n)\n'.format(module)):
        call('pip install {0}'.format(module))
        print('Please restart the program.'.format(module))
        sys.exit()
    else:
        print('This module is important. Please install it.')
        sys.exit()

## Define Functions

def exists(path):
    """Checks to see if a file exists."""
    try:
        with open(path):
            pass
        return True
    except FileNotFoundError:
        return False

def ifttt(action, v1='', v2='', v3=''):
    requests.post('https://maker.ifttt.com/trigger/{0}/with/key/bgj70H05l-3HBc'
    'cRCYvERV'.format(action), data={'value1': v1, 'value2': v2, 'value3': v3})

def get_sheet_corner(workbook_path, sheet_name=None):
    """Returns the column and row of the upper left corner of a spreadsheet.
    
    Just to clarify, if the first cell with data is A1, this script will return
    (1, 1). This is how Excel its numbers."""
    # I have to use x and y because rows and columns get me confused about
    # which way they go
    wb = openpyxl.load_workbook(workbook_path)
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = wb.active
    first_x = 0
    first_y = 0
    corner_found = False
    while corner_found == False:
        if first_x >= 1000:
            raise RuntimeError('No data found for 1000 columns')
        for x in range(first_x, -1, -1):
            y = first_x - x
            temp_cell = ws.cell(row=y + 1, column=x + 1)
            if temp_cell.value:
                return(x + 1, y + 1)
                corner_found = True
        first_x += 1

def read_sheet_column(workbook_path, sheet_name=None, headers=True):
    """Reads the first column in a given sheet.
    
    If headers is True, then loop through all the cells below the upper-left
    corner until a blank space is found. Return a list of all the cells. If a
    cell has a hashtag in the cell to the left of it, do not return that cell.
    """
    corner = get_sheet_corner(workbook_path, sheet_name)
    wb = openpyxl.load_workbook(workbook_path)
    output_cells = []
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = wb.active
    x = corner[0]
    if headers:
        y = corner[1] + 1  # Don't want the headers in the data
    else:
        y = corner[1]
    read_cell = ws.cell(row=y, column=x)
    while read_cell.value:
        read_cell = ws.cell(row=y, column=x)
        if x == 1:
            output_cells.append(read_cell.value)
        else:
            adjacent_cell = ws.cell(row=y-1, column=x)
            if '#' not in str(adjacent_cell.value):
                output_cells.append(read_cell.value)
        if output_cells[-1] == None:
            del output_cells[-1]
        y += 1
    return output_cells

def week(timestamp):
    """Returns the ISO calendar week number of a given timestamp.
    
    Timestamp can be either an integer or a string."""
    return datetime.fromtimestamp(int(timestamp)).isocalendar()[1]

def notify(message):
    """Gives a Pushbullet message."""
    ifttt('notify', v1=message)

def end_script(terminate=True):
    """Ends program."""
    if not isdev:
        input('Press enter to exit')
        sys.exit()
    elif terminate:
        sys.exit()

def error(msg):
    print(msg)
    notify(msg)
    end_script()

def excel_close(file):
    try:
        file.close()
        return True
    except PermissionError:
        return False
        __ = input('Permissions denied! Please close all Excel windows and try'
        ' again.')
        if excel_close(file):
            pass

def rearrange(lst, order):
    """Returns lst but in the order of order.
    
    Indexing starts at 0."""
    return [lst[x] for x in order]

isdev = socket.gethostname() == 'c-laptop'

if isdev:
    print('Developer mode active')
else:
    print('User mode active')
    print(expanduser('~') + '\\Desktop\\stock_signs.txt')

desktop = expanduser('~') + '\\Desktop\\'

## Define Constants

# Backup list of stock signs as of 5/26/2015
backup_signs = ['A', 'ABC', 'ABT', 'ACE', 'ACN', 'ACT', 'ADBE', 'ADI', 'AET', 'AFL', 'AGN', 'AGU', 'AIG', 'ALL', 'ALXN', 'AMGN', 'AMT', 'AMZN', 'APA', 'APC', 'APD', 'AXP', 'AZO', 'BA', 'BAC', 'BAM', 'BAX', 'BBBY', 'BDX', 'BEN', 'BFB', 'BHI', 'BHP', 'BIIB', 'BMY', 'BP', 'BRK-B', 'BUD', 'BWA', 'BXP', 'C', 'CAH', 'CAM', 'CAT', 'CBS', 'CELG', 'CERN', 'CHKP', 'CI', 'CMCSA', 'CME', 'CMG', 'CMI', 'CNQ', 'COF', 'COG', 'COH', 'COST', 'COV', 'CS', 'CSCO', 'CSX', 'CTSH', 'CTXS', 'CVS', 'CVX', 'DAL', 'DD', 'DEO', 'DFS', 'DGX', 'DHR', 'DIS', 'DLPH', 'DOV', 'DTV', 'DVA', 'DVN', 'EBAY', 'ECL', 'EL', 'EMC', 'EMN', 'ENB', 'EOG', 'EPD', 'ESRX', 'ESV', 'ETN', 'F', 'FB', 'FDX', 'FIS', 'FLR', 'GD', 'GE', 'GILD', 'GIS', 'GLW', 'GM', 'GPS', 'GSK', 'GWW', 'HAL', 'HD', 'HES', 'HMC', 'HOG', 'HON', 'HOT', 'HST', 'HSY', 'HUM', 'ICE', 'INTC', 'IP', 'ISRG', 'JCI', 'JNJ', 'JPM', 'KMP', 'KMX', 'KO', 'KR', 'KRFT', 'KSS', 'L', 'LLY', 'LOW', 'LVS', 'LYB', 'M', 'MA', 'MAR', 'MAT', 'MCD', 'MCK', 'MDLZ', 'MDT', 'MET', 'MFC', 'MHFI', 'MMC', 'MO', 'MON', 'MOS', 'MPC', 'MRK', 'MRO', 'MRO', 'MS', 'MSFT', 'MUR', 'MYL', 'NBL', 'NE', 'NEM', 'NKE', 'NLSN', 'NOV','NSC', 'NUE', 'NVS', 'ORCL', 'ORLY', 'OXY', 'PCP', 'PEP', 'PFE', 'PG', 'PH', 'PM', 'PNC', 'PNR', 'PPG', 'PRU', 'PSX', 'PX', 'PXD', 'QCOM', 'QQQ', 'REGN', 'RIO', 'RL', 'ROP', 'ROST', 'RRC', 'RSG', 'SBUX', 'SE', 'SHW', 'SJM', 'SLB', 'SLM', 'SNDK', 'SPG', 'STT', 'STZ', 'SU', 'SWK', 'SYK', 'TCK', 'TEL', 'TJX', 'TM', 'TMO', 'TROW', 'TRV', 'TWC', 'TWX', 'TYC', 'UAL', 'UNH', 'UNP', 'UPS', 'UTX', 'V', 'VFC', 'VIAB', 'VLO', 'VNO', 'VZ', 'WAG', 'WDC', 'WFC', 'WFM', 'WMB', 'WMT', 'WY', 'WYNN', 'YHOO', 'YUM', 'ZMH']

my_backup_signs = ['AAPL', 'INTC', 'GOOG', 'ASDF', 'WMT', 'NFLX']

# Backup list of dates as of 6/24/2015
backup_dates = ['07/02/15', '07/10/15', '07/24/15', '07/31/15', '07/17/15', '08/21/15', '09/18/15', '10/16/15', '11/20/15', '12/18/15']

## Get Signs and Dates

print('Opening files...')

if exists(desktop + 'stock_signs.txt'):
    with open(desktop + 'stock_signs.txt', 'r') as f_readsigns:
        alphabet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ '
        signs = sorted([s.replace('\n', '').replace(' ', '').upper() for s in f_readsigns if s[0] in alphabet])
else:
    with open(desktop + 'stock_signs.txt', 'w') as f_writesigns:
        f_writesigns.write('\n'.join(backup_signs))
        print('stock_signs.txt has been created. Please restart the program.')
        end_script()

if not signs:
    with open(desktop + 'stock_signs.txt', 'w') as f_writesigns:
        f_writesigns.write('\n'.join(backup_signs))

status = {}

try:
    signs = read_sheet_column(desktop + 'stock_signs.xlsx')
except FileNotFoundError:
    write_signs = openpyxl.Workbook()
    write_signs.save(desktop + 'stock_signs.xlsx')
    print('Please go to your desktop and put the dates you want into'
          ' stock_signs.xlsx. Put hash marks in the cells to the left of the'
          ' ones you don\'t want.')
    end_script(terminate=False)
try:
    dates = read_sheet_column(desktop + 'stock_dates.xlsx')
    dates_weeks = [date.isocalendar()[1] for date in dates]
except FileNotFoundError:
    write_dates = openpyxl.Workbook()
    write_dates.save(desktop + 'stock_dates.xlsx')
    print('Please go to your desktop and put the signs you want into'
          ' stock_dates.xlsx. Put hash marks in the cells to the left of the'
          ' ones you don\'t want.')
    end_script(terminate=False)

print('{0} signs, {1} dates'.format(len(signs), len(dates)))

## Miscellaneous Startup

dt = datetime.fromtimestamp(time.time())
date = dt.strftime('%d-%m-%Y')

if not isdev:
    opath = 'C:/Users/Gary/Documents/Option_tables/Option_Model_Files/'
    'OptionReportDirectory/options_report_{0}.xlsx'.format(date)
else:
    opath = 'options_report_{0}.xlsx'.format(date)

try:
    excel = xlsxwriter.Workbook(opath)
except:
    error('Unable to open workbook. Please close it if it is open and try '
        'again.')

start = time.time()

## Download Data

site = 'https://finance.yahoo.com/q/op?s={0}&date={1}'  # .format(sign, date)
first_site = 'https://finance.yahoo.com/q/op?s={0}'  # .format(sign)
left_col = "//div[@id='optionsCallsTable']//tbody/tr"
path_table = "//div[@id='optionsCallsTable']//tbody/tr[{0}]/td/*//text()"
path_last = "//*[@id='yfs_l84_{0}']//text()"  # .format(sign)
path_dates = '//select//@value'


site_2 = 'https://finance.yahoo.com/q/in?s={0}+Industry'  # .format(sign)
paths_info = ['//*[@id="yfi_rt_quote_summary"]/div[1]/div/h2/text()',
    '//tr[1]/td/a/text()', '//tr[2]/td/a/text()']
all_data = {}
errors = []

for sign in signs:
    all_data[sign] = {}
    print('\n{0:{1}} ({2:{3}} of {4})'.format(
            sign, len(max(signs, key=len)) + 1, signs.index(sign) + 1,
            len(str(len(signs))), len(signs)
        ), end='')
    page = requests.get(site_2.format(sign))
    tree = html.fromstring(page.text)
    try:
        all_data[sign]['Info'] = [tree.xpath(paths_info[0])[0]]
    except IndexError:
        print(' Error: stock does not exist.', end='')
        errors.append(sign)
        status[sign] = 'ERROR: stock does not exist'
        continue
    
    try:
        all_data[sign]['Info'].extend([tree.xpath(path)[0] for path in paths_info[1:]])
    except IndexError:
        all_data[sign]['Info'].extend(['EFT', 'EFT'])
    
    page = requests.get(first_site.format(sign))
    tree = html.fromstring(page.text)
    dates_from_site = tree.xpath(path_dates)
    valid_dates = [x for x in dates_from_site if week(x) in dates_weeks]
    if not valid_dates:  # No dates to download, so no dates to do anything with
        status[sign] = 'ERROR: No valid dates'
        continue
    for date in valid_dates:
        all_data[sign][date] = []
        print('.', end='', flush=True)
        page = requests.get(site.format(sign, date))
        tree = html.fromstring(page.text)
        left_data = tree.xpath(left_col)  # So we know how many rows there are
        exists = True
        for row_n in range(len(left_data)):
            temp_row = tree.xpath(path_table.format(row_n + 1))
            try:
                temp_row.insert(0, tree.xpath(path_last.format(sign))[0])
            except IndexError as e:
                exists = False
            if exists:
                all_data[sign][date].append(temp_row)
        if not exists:
            print(' Stock does not exist.', end='')
            break

print()  # Allow printing of the last line

download_end = time.time()
try:
    print(
        'Download completed in {0:.2f} seconds (average {0:.2f} seconds per'
        ' stock)'.format(
            download_end - start,
            (download_end - start) / len(signs)
    )
except ZeroDivisionError:
    error(
        'No stock signs found. Please enter them into stock_signs.txt on your'
        ' desktop and try again.'
    )

## Format Data

formats = [
    'str', 'str', 'str', 'str', 'float',
    'str', 'str_f', 'str', 'float', 'float',
    'float', 'int', 'int', 'float', 'float_f',
    'int_f', 'int_f', 'int_f', 'float_f', 'percent_f',
    'percent_f', 'float_f', 'percent_f', 'percent_f', 'str_f'
]
headers = [
    'co_symbol', 'company', 'industry', 'sector', 'Last', 'Option', 'exp_date',
    'Call', 'Strike', 'Bid', 'Ask', 'Open interest', 'Vol', 'Last',
    datetime.now().strftime('%m/%d/%y'), 'days', '60000', ' $invested',
    '$prem', ' prem%', 'annPrem%', ' MaxRet', ' Max%', 'annMax%', '10%'
]
formulas = [
    '=IF(J{n}<F{n},(J{n}-F{n})+K{n},K{n})', '=H{n}-P$6',
    '=ROUND(R$6/((F{n}-0)*100),0)', '=100*R{n}*(F{n}-0)', '=100*P{n}*R{n}',
    '=(T{n}/S{n})*100', '=(365/Q{n})*U{n}',
    '=IF(J{n}>F{n},(100*R{n}*(J{n}-F{n}))+T{n},T{n})', '=(W{n}/S{n})*100',
    '=((365/Q{n})*X{n})*100', '=IF((ABS(J{n}-F{n})/J{n})<Z$6,"NTM","")'
]

data = []

for sign in all_data:
    for date in all_data[sign]:
        if date != 'Info':
            for r in all_data[sign][date][:]:
                # human-readable date = hrd
                try:
                    hrd_lst = [r[2][-15:-9][x:x + 2] for x in range(0, 6, 2)]
                except IndexError as ie:
                    raise IndexError(ie.args, r) from ie
                hrd_str = '/'.join((hrd_lst[1], hrd_lst[2], hrd_lst[0]))
                try:
                    row = ([sign] + all_data[sign]['Info'][0:3] + 
                        rearrange(r, [0, 2]) + [hrd_str, 'C'] + 
                        rearrange(r, [1, 4, 5, 9, 8, 3]) + formulas)
                except IndexError as ie:
                    raise IndexError(row) from ie
                data.append(row)

fails = {'Signs': [], 'Dates': []}
for sign in all_data:
    for date in all_data[sign]:
        pass
# Check that everything that's supposed to be the same length is

in_data = {}
for sign in signs:
    in_data[sign] = False
    for row in data:
        for cell in row:
            if sign in cell:
                in_data[sign] = True

assert data

no_data = [sign for sign in in_data if not in_data[sign]]
if errors != []:
    print('The following stocks failed to download: {0}.'.format(', '.join(errors)))
if [x for x in no_data if x not in errors] != []:
    print('The following stocks returned no data: {0}.'.format(', '.join([x for x in no_data if x not in errors])))

try:
    if len(formats) != len(headers) or len(headers) != len(data[0]):
        error('The "formats" list, "headers" list, and rows in the data are'
              ' not all the same length!')
except IndexError as e:
    raise IndexError from e

for row in data:
    for i, cell in enumerate(row):
        if '_f' in formats[i]:
            row[i] = str(row[i])
        elif 'percent' in formats[i]:
            row[i] = float(row[i].replace('%', ''))
        else:
            try:
                row[i] = eval('{0}(row[i])'.format(formats[i].replace('_f', '')))
            except ValueError:
                if '-' in row[i]:
                    row[i] = str(row[i])

data_sector = {r[2]: [] for r in data}
for r in data:
    data_sector[r[2]].append(r)


## Output Data

write_start = time.time()


sheets = {s: excel.add_worksheet(s) for s in data_sector}
def_r_offset, def_c_offset = 5, 1  # Defaults

pt = excel.add_format({'num_format': '#,##0.00\%'})  # percent
ft = excel.add_format({'num_format': '#,##0.00'})  # float
it = excel.add_format({'num_format': '#,##0'})  # int
sr = excel.add_format({})  # str
fa = excel.add_format({})  # formula

print('Writing data...', end='')

for sector in data_sector:
    r_offset, c_offset = def_r_offset, def_c_offset
    for i, header in enumerate(headers):
            sheets[sector].write(r_offset, i + c_offset, header)

    r_offset += 1

    for r, row in enumerate(data_sector[sector]):
        for c, cell in enumerate(row):
            if '_f' in formats[c]:
                sheets[sector].write(r + r_offset, c + c_offset, cell.format(n=str(r + r_offset + 1)), eval(formats[c][0] + formats[c][-3]))
            else:
                sheets[sector].write(r + r_offset, c + c_offset, cell, eval(formats[c][0] + formats[c][-1]))

excel_close(excel)

## Finish Up

end = time.time()
print(' Completed in {0:.2f} seconds'.format(end - write_start))
print('Script completed in {0:.2f} seconds'.format(end - start))

notify('Your script has just been run on {0}, taking a total of {1} seconds to download and write {2} stocks and {3} dates.'.format(socket.gethostname(), end - start, len(signs), len(dates)))

# requests.post('https://maker.ifttt.com/trigger/script_logged/with/key/bgj70H05l-3HBccRCYvERV', data={'value1': '{0} ||| {1} ||| {2} ||| {3}'.format(socket.gethostname(), len(signs), len(dates), end - start)})

ifttt('script_logged', v1='{0} ||| {1} ||| {2} ||| {3}'.format(socket.gethostname(), len(signs), len(dates), end - start))

if 'y' in input('Would you like to open the file in Excel? (y/n) ').lower():
    try:
        os.startfile(opath)
    except OSError:
        print('Unable to open Excel. The file is called {0}.'.format(path.split('/')[-1]))
        
end_script(terminate=False)