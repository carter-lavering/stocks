## To-Do List

# [X] Separate sheets for sectors
# [X] Add easier way to input dates
# [X] Add separate files for dates and signs
# [ ] Automated date validity check

## Import Modules


from datetime import datetimeimport time
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
        with open(path) as test_file:
            pass
        return True
    except FileNotFoundError:
        return False

def ifttt(action, v1='', v2='', v3=''):
    requests.post('https://maker.ifttt.com/trigger/{0}/with/key/bgj70H05l-3HBccRCYvERV'.format(action), data={'value1': v1, 'value2': v2, 'value3': v3})

def notify(message):
    """Gives a Pushbullet message."""
    ifttt('notify', v1=message)

def closest_num(n, lst):
    """Returns the closest number to n in lst.
    
    Assumes all items in the list are numbers."""
    lst_by_diffs = {abs(n - x): x for x in lst}
    diffs = sorted([x for x in lst_by_diffs])
    return lst_by_diffs[diffs[0]]

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
    except PermissionError:
        __ = input('Permissions denied! Please close all Excel windows and try again.')
        excel_close(file)

def rearrange(lst, order):
    """Returns lst but in the order of, well, order."""
    return [lst[x] for x in order]

def empty_list(lst):
    empty = True
    for x in lst:
        if type(x) == list:
            empty = empty_list(x)
        else:
            empty = False
    return empty

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

if signs == []:
    with open(desktop + 'stock_signs.txt', 'w') as f_writesigns:
        f_writesigns.write('\n'.join(backup_signs))

print(fill('{0} signs:\n{1}'.format(len(signs), ', '.join(signs)), 80))

if exists(desktop + 'stock_dates.txt'):
    with open(desktop + 'stock_dates.txt', 'r') as f_readdates:
        numbers = '0123456789 '
        dates_hr = sorted([d.replace('\n', '').replace(' ', '').upper() for d in f_readdates if d[0] in numbers])
        dates = [str(int(time.mktime(time.strptime(d, '%m/%d/%y'))) - time.timezone) for d in dates_hr]
else:
    with open(desktop + 'stock_dates.txt', 'w') as f_writedates:
        f_writedates.write('\n'.join(backup_dates))
    print('stock_dates.txt has been created. Please restart the program.')
    end_script()

print('{0} dates:'.format(len(dates_hr)), ', '.join(dates_hr))

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
        all_data[sign]['Info'] = [tree.xpath(path)[0] for path in paths_info]
    except IndexError:
        print(' Error: stock does not exist.', end='')
        errors.append(sign)
    else:
        valid = 0
        page = requests.get(first_site.format(sign))
        tree = html.fromstring(page.text)
        dates_from_site = tree.xpath(path_dates)
        for date in [x for x in dates if x in dates_from_site]:
            all_data[sign][date] = []
            print('.', end='')
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
    print('Download completed in {0:.2f} seconds (average {0:.2f} seconds per stock)'.format(download_end - start, (download_end - start) / len(signs)))
except ZeroDivisionError:
    error('No stock signs found. Please enter them into stock_signs.txt on your desktop and try again.')

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
if no_data != []:
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

data_sector = {r[3]: [] for r in data}
for r in data:
    data_sector[r[3]].append(r)


## Output Data

write_start = time.time()

sheets = {s: excel.add_worksheet(s) for s in data_sector}
def_r_offset, def_c_offset = 5, 1  # Defaults

pt = excel.add_format({'num_format': '#,##0.00\%'})
ft = excel.add_format({'num_format': '#,##0.00'})
it = excel.add_format({'num_format': '#,##0'})
sr = excel.add_format({})
fa = excel.add_format({})

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
if errors != []:
    print('The following stocks failed to download: {0}.'.format(', '.join(errors)))
else:
    print('All stocks downloaded successfully.')

try:
    os.startfile(opath)
except OSError:
    print('Unable to open Excel. The file is called {0}.'.format(path.split('/')[-1]))

notify('Your script has just been run on {0}, taking a total of {1} seconds to download and write {2} stocks and {3} dates.'.format(socket.gethostname(), end - start, len(signs), len(dates)))

requests.post('https://maker.ifttt.com/trigger/script_logged/with/key/bgj70H05l-3HBccRCYvERV', data={'value1': '{0} ||| {1} ||| {2} ||| {3}'.format(socket.gethostname(), len(signs), len(dates), end - start)})

end_script(terminate=False)