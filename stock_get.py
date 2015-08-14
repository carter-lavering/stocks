#!/usr/bin/env python
# coding=utf-8

# \_\_\_\_\_    \_\_\_                \_\_\_\_      \_\_\_
#     \_      \_      \_              \_      \_  \_      \_
#      \_      \_      \_  \_\_\_\_\_  \_      \_  \_      \_
#       \_      \_      \_              \_      \_  \_      \_
#        \_        \_\_\_                \_\_\_\_      \_\_\_

# [X] Separate sheets for sectors
# [X] Add easier way to input dates
# [X] Add separate files for dates and symbols
# [ ] Automated date validity check

# \_\_\_\_\_  \_      \_  \_\_\_\_      \_\_\_    \_\_\_\_    \_\_\_\_\_
#      \_      \_\_  \_\_  \_      \_  \_      \_  \_      \_      \_
#       \_      \_  \_  \_  \_\_\_\_    \_      \_  \_\_\_\_        \_
#        \_      \_      \_  \_          \_      \_  \_    \_        \_
#     \_\_\_\_\_  \_      \_  \_            \_\_\_    \_      \_      \_

import socket
import sys
import time
from datetime import datetime
from os.path import expanduser

import openpyxl
import requests
from lxml import html


# \_\_\_\_    \_\_\_\_\_  \_\_\_\_\_  \_\_\_\_\_  \_      \_  \_\_\_\_
#  \_      \_  \_          \_              \_      \_\_    \_  \_
#   \_      \_  \_\_\_\_\_  \_\_\_\_\_      \_      \_  \_  \_  \_\_\_\_
#    \_      \_  \_          \_              \_      \_    \_\_  \_
#     \_\_\_\_    \_\_\_\_\_  \_          \_\_\_\_\_  \_      \_  \_\_\_\_


def ifttt(action, v1='', v2='', v3=''):
    requests.post(
        'https://maker.ifttt.com/trigger/{0}/with/key/bgj70H05l-3HBccRCYvERV'
        .format(action), data={'value1': v1, 'value2': v2, 'value3': v3})


def get_sheet_corner(workbook_path, sheet_name=None):
    """Returns the column and row of the upper left corner of a spreadsheet.

    Indexing starts at 1, so A1 is (1, 1), not (0, 0)."""
    # I have to use x and y because rows and columns get me confused about
    # which way they go
    wb = openpyxl.load_workbook(workbook_path)
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = wb.active
    first_x = 0
    corner_found = False
    while not corner_found:
        if first_x >= 1000:
            raise RuntimeError('No data found for 1000 columns')
        for x in range(first_x, -1, -1):
            y = first_x - x
            temp_cell = ws.cell(row=y + 1, column=x + 1)
            if temp_cell.value:
                return (x + 1, y + 1)
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
        if output_cells[-1] is None:
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
        input('Permissions denied! Please close all Excel windows and try'
              ' again.')
        if excel_close(file):
            pass


def rearrange(lst, order):
    """Returns lst but in the order of order.

    Indexing starts at 0."""
    return [lst[x] for x in order]


# \_\_\_\_\_  \_\_\_\_\_  \_\_\_\_\_  \_\_\_\_    \_\_\_\_\_
#  \_              \_      \_      \_  \_      \_      \_
#   \_\_\_\_\_      \_      \_\_\_\_\_  \_\_\_\_        \_
#            \_      \_      \_      \_  \_    \_        \_
#     \_\_\_\_\_      \_      \_      \_  \_      \_      \_

version = '1.0.1'
print('Stock data downloader version {0}'.format(version))
isdev = socket.gethostname() == 'c-laptop'
if isdev:
    print('Developer mode active')
else:
    print('User mode active')
    print(expanduser('~') + '\\Desktop\\stock_signs.txt')
desktop = expanduser('~') + '\\Desktop\\'

print('Opening files...')
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
          " ones you don't want.")
    end_script(terminate=False)

assert signs
assert dates

print('{0} signs, {1} dates'.format(len(signs), len(dates)))

dt = datetime.fromtimestamp(time.time())
date = dt.strftime('%d-%m-%Y')

if not isdev:
    output_path = (
        'C:/Users/Gary/Documents/Option_tables/Option_Model_Files/'
        'OptionReportDirectory/options_report_{0}.xlsx'.format(date)
    )
else:
    output_path = 'options_report_{0}.xlsx'.format(date)

test_save_location = openpyxl.Workbook()
test_save_location.save(output_path)

start = time.time()

# Download Data

site = 'https://finance.yahoo.com/q/op?s={0}&date={1}'  # .format(sign, date)
first_site = 'https://finance.yahoo.com/q/op?s={0}'  # .format(sign)
left_col = "//div[@id='optionsCallsTable']//tbody/tr"
path_table = "//div[@id='optionsCallsTable']//tbody/tr[{0}]/td/*//text()"
path_last = "//*[@id='yfs_l84_{0}']//text()"  # .format(sign)
path_dates = '//select//@value'


site_2 = 'https://finance.yahoo.com/q/in?s={0}+Industry'  # .format(sign)
paths_info = ['//*[@id="yfi_rt_quote_summary"]/div[1]/div/h2/text()',
              '//tr[1]/td/a/text()',
              '//tr[2]/td/a/text()']
all_data = {}
errors = []

for sign in signs:
    all_data[sign] = {}
    print('\n{0:{1}} ({2:{3}} of {4})'.format(sign,
                                              len(max(signs, key=len)),
                                              signs.index(sign) + 1,
                                              len(str(len(signs))),
                                              len(signs)), end='')
    page = requests.get(site_2.format(sign))
    tree = html.fromstring(page.text)
    try:
        all_data[sign]['Info'] = [tree.xpath(paths_info[0])[0]]
    except IndexError:
        print(' Error: stock does not exist.', end='')
        errors.append(sign)
        continue

    try:
        all_data[sign]['Info'].extend(
            [tree.xpath(path)[0] for path in paths_info[1:]]
        )
    except IndexError:
        all_data[sign]['Info'].extend(['EFT', 'EFT'])

    page = requests.get(first_site.format(sign))
    tree = html.fromstring(page.text)
    dates_from_site = tree.xpath(path_dates)
    valid_dates = [x for x in dates_from_site if week(x) in dates_weeks]
    if not valid_dates:
        print(' Error: No valid dates.', end='')
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
        'Download completed in {0:.2f} seconds (average {1:.2f} seconds per'
        ' stock)'.format(
            download_end - start,
            (download_end - start) / len(signs)
        )
    )
except ZeroDivisionError:
    error(
        'No stock signs found. Please enter them into stock_signs.txt on your'
        ' desktop and try again.'
    )

# \_\_\_\_\_    \_\_\_    \_\_\_\_    \_      \_  \_\_\_\_\_  \_\_\_\_\_
#  \_          \_      \_  \_      \_  \_\_  \_\_  \_      \_      \_
#   \_\_\_\_    \_      \_  \_\_\_\_    \_  \_  \_  \_\_\_\_\_      \_
#    \_          \_      \_  \_    \_    \_      \_  \_      \_      \_
#     \_            \_\_\_    \_      \_  \_      \_  \_      \_      \_

headers = ['co_symbol', 'company', 'sector', 'industry', 'Last', 'exp_date',
           'Strike', 'Bid', 'Ask', 'Vol', 'Last',
           datetime.now().date(), 'days', '60000', ' $invested',
           '$prem', ' prem%', 'annPrem%', ' MaxRet', ' Max%', 'annMax%', '10%']

formulas = [
    '=IF(H{n}<F{n},(H{n}-F{n})+I{n},I{n})',
    '=G{n}-M$6',
    '=ROUND(O$6/((F{n}-0)*100),0)',
    '=100*O{n}*(F{n}-0)',
    '=100*M{n}*O{n}',
    '=(Q{n}/P{n})*100',
    '=(365/N{n})*R{n}',
    '=100*M{n}*O{n}',
    '=(T{n}/P{n})*100',
    '=((365/N{n})*U{n})*100',
    '=IF((ABS(H{n}-F{n})/H{n})<W$6,"NTM","")'
]

data = []

for sign in all_data:
    for date in [date for date in all_data[sign] if date != 'Info']:
        for r in all_data[sign][date][:]:
            # human-readable date = hrd
            # computer-readable date = crd
            hrd_lst = [r[2][-15:-9][x:x + 2] for x in range(0, 6, 2)]
            # Don't delete the extra parentheses, .join() only takes one
            # argument
            hrd_str = '/'.join((hrd_lst[1], hrd_lst[2], '20' + hrd_lst[0]))
            crd = datetime.strptime(hrd_str, '%m/%d/%Y').date()
            if hrd_str[0] == '0':  # No zeros at the beginning
                hrd_str = hrd_str[1:]
            row = ([sign] +
                   all_data[sign]['Info'][:3] +
                   [r[0]] +
                   [crd] +
                   rearrange(r, [1, 4, 5, 8, 3]) +
                   formulas)
            data.append(row)

in_data = {}
for sign in signs:
    in_data[sign] = False
    for row in data:
        for cell in row:
            try:
                if sign in cell:
                    in_data[sign] = True
                break
            except TypeError:
                pass

# Make sure it actually has things in it
assert data
# Make sure everything's the same length
try:
    assert len(headers) == len(data[0])
except AssertionError as e:
    raise AssertionError(e.args, len(headers), len(data[0]))

no_data = [sign for sign in in_data if not in_data[sign] if sign not in errors]
if errors != []:
    print('The following stocks failed to download: {0}.'.format(
        ', '.join(errors)
    ))
if [x for x in no_data if x not in errors] != []:
    print('The following stocks returned no data: {0}.'.format(
        ', '.join(no_data)
    ))

data_sector = {r[2]: [] for r in data}
for r in data:
    data_sector[r[2]].append(r)


# \_      \_  \_\_\_\_    \_\_\_\_\_  \_\_\_\_\_  \_\_\_\_\_
#  \_      \_  \_      \_      \_          \_      \_
#   \_      \_  \_\_\_\_        \_          \_      \_\_\_\_\_
#    \_  \_  \_  \_    \_        \_          \_      \_
#     \_\_  \_\_  \_      \_  \_\_\_\_\_      \_      \_\_\_\_\_

write_start = time.time()

print('Writing data...', end=' ')
workbook = openpyxl.Workbook(guess_types=True)
master_sheet = workbook.active
master_sheet.title = 'Master'
sheet_names = sorted([sheet for sheet in data_sector])  # Alphabetical
for name in sheet_names:
    temp_sheet = workbook.create_sheet()
    temp_sheet.title = name

for sheet in workbook:
    for i, header in enumerate(headers):
        sheet.cell(row=6, column=2+i).value = header

for sector in data_sector:
    for r, row in enumerate(data_sector[sector]):
        for c, cell in enumerate(row):
            try:
                workbook.get_sheet_by_name(sector).cell(
                    row=7+r,
                    column=2+c
                ).value = cell.format(n=7+r)
            except AttributeError:
                workbook.get_sheet_by_name(sector).cell(
                    row=7+r,
                    column=2+c
                ).value = cell

r = 0  # Write all the data, so it can't restart counting every time
for sector in data_sector:
    for row in data_sector[sector]:
        for c, cell in enumerate(row):
            try:
                master_sheet.cell(row=7+r,
                                  column=2+c).value = cell.format(n=7+r)
            except AttributeError:
                master_sheet.cell(
                    row=7+r,
                    column=2+c
                ).value = cell
        r += 1


workbook.save(output_path)


# Finish Up

end = time.time()
print('Completed in {0:.2f} seconds'.format(end - write_start))
print('Script completed in {0:.2f} seconds'.format(end - start))

notify(
    'Your script has just been run on {0}, taking a total of {1} seconds to'
    ' download and write {2} stocks and {3} dates.'.format(
        socket.gethostname(),
        round(end - start, 2),
        len(signs),
        len(dates)
    )
)

ifttt('script_logged', v1='{0} ||| {1} ||| {2} ||| {3}'.format(
    socket.gethostname(),
    len(signs),
    len(dates),
    end - start
))

# if 'y' in input('Would you like to open the file in Excel? (y/n) ').lower():
#     try:
#         os.startfile(output_path)
#     except OSError:
#         print('Unable to open Excel. The file is called {0}.'.format(
#             path.split('/')[-1])
#         )

end_script(terminate=False)
