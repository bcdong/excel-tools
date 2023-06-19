#!/usr/bin/python3

import argparse
import xlwings as xw
from datetime import datetime

_DATE_COL_NAMES = ['创建时间', '修改时间', '创建日期', '修改日期']


def fmt_date_cols(sht):
    ncols = sht.used_range.last_cell.column
    head_vals = sht[0, 0:ncols].value
    if not head_vals:
        return
    for col_id, col_name in enumerate(head_vals):
        if col_name not in _DATE_COL_NAMES:
            continue
        col_letter = xw.utils.col_name(col_id + 1)
        print('Setting col=[{}] with name=[{}] format to yyyy-mm-dd ...'.format(col_letter, col_name))
        sht.range('{0}:{0}'.format(col_letter)).number_format = 'yyyy-mm-dd'


def copy_sheets(app, ibook, ofile, with_format, should_fmt_date):
    all_sheets = ibook.sheet_names
    print('All existing sheets: {}'.format(all_sheets))
    out_sheets = input("Please input the sheet names (split by english ,) to extract, or input 'q' to exit:\n")
    if out_sheets.strip() == 'q':
        return
    out_sheets = out_sheets.strip().split(',')
    print('Creating output file...')
    obook = app.books.add()
    out_idx = 0
    print('Copying sheets to output file...')
    for sheet_name in out_sheets:
        sheet_name = sheet_name.strip()
        if sheet_name not in all_sheets:
            print('Invalid sheet name: {}. Continue to copy next sheet.'.format(sheet_name))
            continue
        print('===== Copying sheet: {}, start at {}'.format(sheet_name, datetime.now().strftime('%H:%M:%S')))
        if with_format:
            ibook.sheets[sheet_name].copy(after=obook.sheets[out_idx])
        else:
            obook.sheets.add(name=sheet_name, after=obook.sheets[out_idx])
            output_data = ibook.sheets[sheet_name][0, 0].expand('table').value
            obook.sheets[sheet_name][0, 0].expand('table').value = output_data
        if should_fmt_date:
            fmt_date_cols(obook.sheets[sheet_name])
        out_idx += 1
        print('===== Done copy sheet: {}, end at {}'.format(sheet_name, datetime.now().strftime('%H:%M:%S')))

    if out_idx > 0:
        obook.sheets[0].delete()
    obook.save(ofile)
    obook.close()
    print('Copy {} sheets done!'.format(out_idx))


def process_excel(ifile, ofile, with_format, fmt_date):
    if not ofile.endswith('.xlsx'):
        print('Error: output file name must ends with .xlsx!')
        return
    app=xw.App(visible=True, add_book=False)
    app.display_alerts=False
    app.screen_updating=False
    print('Opening input file...')
    ibook = app.books.open(ifile)
    copy_sheets(app, ibook, ofile, with_format, fmt_date)
    ibook.close()
    app.quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Extract some sheets from one excel and output selected sheets into a new excel')
    parser.add_argument('--ifile', action='store', required=True,  help='the input file')
    parser.add_argument('--ofile', action='store', required=True,  help='the output file')
    parser.add_argument('--with_format', action='store', required=True, choices=['yes', 'no'],  help='whether to preserve formats')
    parser.add_argument('--fmt_date', action='store', required=True, choices=['yes', 'no'],  help='whether to format datetime cols')
    args = parser.parse_args()

    process_excel(args.ifile, args.ofile, args.with_format == 'yes', args.fmt_date == 'yes')
