#!/usr/bin/python3

import argparse
import xlwings as xw
from datetime import datetime

_SUPPORT_TIME_FMT = ['%Y-%m-%d', '%m/%d/%Y']

def parse_datetime(date_str):
    print('Parse time: {}'.format(date_str))
    for fmt in _SUPPORT_TIME_FMT:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            pass
    raise ValueError('date_str=[{}] does not match any fmt'.format(date_str))


def format_col_time(sheet, col_id):
    need_update = False
    col_vals = sheet[1, col_id].expand('down').value
    for i in range(len(col_vals)):
        print('fmt datetime: {}, type={}'.format(col_vals[i], type(col_vals[i])))
        if type(col_vals[i]) is str:
            col_vals[i] = parse_datetime(col_vals[i])
            need_update = True
        elif type(col_vals[i]) is datetime:
            pass
        else:
            raise ValueError('Unsupported type of datetime: {}'.format(type(col_vals[i])))
    if need_update:
        sheet[1, col_id].options(transpose=True).value = col_vals


def sort_sheet(sht, col_id):
    nrows = sht.used_range.last_cell.row
    ncols = sht.used_range.last_cell.column
    # sht[1:nrows, 0:ncols] means sorting data excluding first header row
    # Key1=sht[0, col_id].api means sort according to col_id
    # Order1=1 means ascending, 2 means descending
    # Orientation=1 means sort in columns but not rows
    sht[1:nrows, 0:ncols].api.Sort(Key1=sht[0, col_id].api, Order1=1, Orientation=1)

def format_and_sort_sheet(sheet, col_name):
    print('===== Sorting in sheet: [{}] ====='.format(sheet.name))
    ncols = sheet.used_range.last_cell.column
    head_vals = sheet[0, 0:ncols].value
    if not head_vals:
        print('Empty header in sheet [{}], skip it'.format(sheet.name))
        return
    try:
        col_id = head_vals.index(col_name)
    except ValueError:
        col_id = -1
    if col_id < 0:
        print('column [{}] not exist in sheet [{}], skip it'.format(col_name, sheet.name))
        return
    format_col_time(sheet, col_id)
    sort_sheet(sheet, col_id)


def sort_book(ibook, col_name):
    for sht in ibook.sheets:
        format_and_sort_sheet(sht, col_name)


def process_excel(ifile):
    app=xw.App(visible=True, add_book=False)
    app.display_alerts=False
    app.screen_updating=False
    print('Opening input file...')
    ibook = app.books.open(ifile)

    while True:
        col_name = input("Please input the col_name to sort, or input 'q' to exit:\n")
        if col_name.rstrip() == 'q':
            break
        col_name = col_name.rstrip()
        sort_book(ibook, col_name)

    print('Saving file...')
    ibook.save()
    ibook.close()
    app.quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Sort according to supplied column in all sheets')
    parser.add_argument('--ifile', action='store', required=True,  help='the input file')
    args = parser.parse_args()
    process_excel(args.ifile)
