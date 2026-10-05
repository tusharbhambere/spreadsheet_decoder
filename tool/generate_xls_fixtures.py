#!/usr/bin/env python3
"""Generates the .xls fixtures used by the tests (requires `pip install xlwt`).

    python3 tool/generate_xls_fixtures.py test/files
"""
import datetime
import os
import sys

import xlwt

out = sys.argv[1] if len(sys.argv) > 1 else 'test/files'


def style(fmt):
    return xlwt.easyxf(num_format_str=fmt)


# Same content as test.xlsx
wb = xlwt.Workbook(encoding='utf-8')
one = wb.add_sheet('ONE')
for r, row in enumerate([
    ['A', 'B', 'C'], [1, 2, 3], [4, 5, 6], [7, 8, 9], [12, 15, 18],
    [3, 3, 3], [3, 3, 3], [3, 3, 3], [None, None, None], [6, 7, 8],
    [6, 7, 8], [6, 7, 8],
]):
    for c, v in enumerate(row):
        if v is not None:
            one.write(r, c, v)
two = wb.add_sheet('TWO')
for r, row in enumerate([
    ['X', 'Y', 'Z'], [10, 11, 12], [13, None, 15], [16, 17, 18],
    ["&é'(§è!çà)-", '\u00a0"«A»"', '<>'],
]):
    for c, v in enumerate(row):
        if v is not None:
            two.write(r, c, v)
three = wb.add_sheet('THREE')
for r, row in enumerate([
    ['P', 'Q', 'R'], [100, 101, 102], [103, 104, 105], [106, 107, 108],
    ['A', 'B\nC', 'D\nE\nF'],
]):
    for c, v in enumerate(row):
        three.write(r, c, v)
wb.add_sheet('EMPTY')
wb.save(os.path.join(out, 'test.xls'))

# Value types, dates and awkward layouts
wb = xlwt.Workbook(encoding='utf-8')
ws = wb.add_sheet('Types')
ws.write(0, 0, 'text')
ws.write(0, 1, 42)
ws.write(0, 2, -7)
ws.write(0, 3, 3.14159)
ws.write(0, 4, -1500.99)
ws.write(0, 5, 0.01)          # RK with the "divided by 100" flag
ws.write(0, 6, 1234567890123)  # too large for an RK integer
ws.write(0, 7, True)
ws.write(0, 8, False)
ws.write(1, 0, '日本語と漢字')   # UTF-16 string
ws.write(1, 1, 'Ünïcödé àçcents')
ws.write(1, 2, 'line1\nline2')
ws.write(1, 3, '😀 emoji')       # surrogate pair
# Sparse layout: gaps in rows and columns
ws.write(4, 2, 'C5')
ws.write(7, 5, 'F8')
d = datetime.datetime(2008, 7, 21, 13, 45, 30)
ws.write(9, 0, d, style('M/D/YY'))             # built-in id 14
ws.write(9, 1, d, style('D-MMM-YY'))           # built-in id 15
ws.write(9, 2, d, style('DD/MM/YYYY'))         # custom date format
ws.write(9, 3, d, style('YYYY-MM-DD HH:MM'))   # custom date time format
ws.write(10, 0, datetime.time(13, 45, 30), style('h:mm:ss'))  # built-in 21
ws.write(10, 1, 0.5, style('0.00%'))           # number, not a date
ws.write(10, 2, 1234.5, style('#,##0.00'))
ws.write(10, 3, 1, style('General'))
wb.save(os.path.join(out, 'types.xls'))

# 1904 date system
wb = xlwt.Workbook(encoding='utf-8')
wb.dates_1904 = 1
ws = wb.add_sheet('Dates1904')
ws.write(0, 0, datetime.datetime(2008, 7, 21), style('M/D/YY'))
ws.write(0, 1, datetime.datetime(1904, 1, 1), style('M/D/YY'))
wb.save(os.path.join(out, 'date1904.xls'))

# Many strings: the shared string table spans several CONTINUE records, with
# strings split inside their characters and a switch between 8 and 16 bit.
wb = xlwt.Workbook(encoding='utf-8')
ws = wb.add_sheet('Strings')
for i in range(2500):
    text = 'row %d ' % i + ('é' * (i % 7))
    if i % 11 == 0:
        text += ' 日本語'
    ws.write(i % 1000, i // 1000, text)
ws.write(0, 3, 'x' * 9000)           # longer than one record
ws.write(1, 3, '日' * 5000)          # wide string longer than one record
wb.save(os.path.join(out, 'strings.xls'))

# Several sheets in one workbook, one sheet name with a space and accents
wb = xlwt.Workbook(encoding='utf-8')
a = wb.add_sheet('First sheet')
a.write(0, 0, 'a')
b = wb.add_sheet('Éte')
b.write(0, 0, 'b')
b.write(1, 1, 2)
wb.save(os.path.join(out, 'sheets.xls'))
