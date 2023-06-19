#!/usr/bin/python3

import argparse
import xlwings as xw
from datetime import datetime

_SUPPORT_TIME_FMT = ['%Y-%m-%d']

def parse_datetime(date_str):
    for fmt in _SUPPORT_TIME_FMT:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            pass
    #raise ValueError('date_str=[{}] does not match any fmt'.format(date_str))
    return date_str  # return the unrecongnized value as str


def format_col_time(sheet, col_id, col_name):
    nrows = sheet.used_range.last_cell.row
    col_vals = sheet[1:nrows, col_id].value
    num_str = 0
    for i in range(len(col_vals)):
        if col_vals[i] is None:
            # ignore empty value
            pass
        elif type(col_vals[i]) is datetime:
            pass
        elif type(col_vals[i]) is str:
            old_val = col_vals[i]
            col_vals[i] = parse_datetime(col_vals[i])
            if type(col_vals[i]) is datetime:
                print('format datetime from str, row={}, val={}'.format(i + 2, old_val))
                num_str += 1
        else:
            raise ValueError('Unsupported type of datetime: {}'.format(type(col_vals[i])))
    col_letter = xw.utils.col_name(col_id + 1)
    # print('insert column {}'.format(col_name))
    # sheet.range('{0}:{0}'.format(col_letter)).api.Insert()  # insert new col left to original col
    print('Total number of formated datetime: {}'.format(num_str))
    print('Setting col number format to yyyy-mm-dd ...')
    sheet.range('{0}:{0}'.format(col_letter)).number_format = 'yyyy-mm-dd'
    # print('Setting col values...')
    sheet[1, col_id].options(transpose=True).value = col_vals
    # print('Done set col values...')


def sort_sheet(sht, col_id):
    nrows = sht.used_range.last_cell.row
    ncols = sht.used_range.last_cell.column
    # sht[1:nrows, 0:ncols] means sorting data excluding first header row
    # Key1=sht[0, col_id].api means sort according to col_id
    # Order1=1 means ascending, 2 means descending
    # Orientation=1 means sort in columns but not rows
    sht[1:nrows, 0:ncols].api.Sort(Key1=sht[0, col_id].api, Order1=1, Orientation=1)

def format_and_sort_sheet(sheet, col_name, fmt_date):
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
        # print('column [{}] not exist in sheet [{}], skip it'.format(col_name, sheet.name))
        return
    print('===== Sorting sheet: [{}], start at {}'.format(sheet.name, datetime.now().strftime('%H:%M:%S')))
    if fmt_date:
        format_col_time(sheet, col_id, col_name)
    sort_sheet(sheet, col_id)
    print('===== Finish sorting sheet: [{}], end at {}'.format(sheet.name, datetime.now().strftime('%H:%M:%S')))


def sort_book(ibook, col_name, fmt_date):
    for sht in ibook.sheets:
        format_and_sort_sheet(sht, col_name, fmt_date)


def process_excel(ifile, fmt_date):
    app=xw.App(visible=True, add_book=False)
    app.display_alerts=False
    app.screen_updating=False
    print('Opening input file...')
    ibook = app.books.open(ifile)

    while True:
        col_name = input("Please input the col_name to sort, or input 'q' to save and exit:\n")
        if col_name.strip() == 'q':
            break
        col_name = col_name.strip()
        sort_book(ibook, col_name, fmt_date)

    print('Saving file...')
    ibook.save()
    ibook.close()
    app.quit()
    print('Done!')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Sort according to supplied column in all sheets')
    parser.add_argument('--ifile', action='store', required=True,  help='the input file')
    parser.add_argument('--fmt_date', action='store', required=True, choices=['yes', 'no'],  help='whether to format datetime. Very slow')
    args = parser.parse_args()
    process_excel(args.ifile, args.fmt_date == 'yes')
