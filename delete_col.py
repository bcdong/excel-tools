#!/usr/bin/python3

import argparse
import xlwings as xw

def del_cols_in_one_sheet(sheet, col_names):
    print('Deleting columns in sheet: {}'.format(sheet.name))
    ncols = sheet[0: 0].expand("right").last_cell.column
    head_vals = sheet[0, 0:ncols].value
    for col_name in col_names:
        try:
            col_id = head_vals.index(col_name)
        except ValueError:
            col_id = -1
        if col_id < 0:
            print('column {} not exist in sheet {}, skip it'.format(col_name, sheet.name))
        else:
            col_letter = xw.utils.col_name(col_id + 1)  # xw.utils.col_name needs col_id starts from 1 but not 0
            sheet.range('{0}:{0}'.format(col_letter)).api.Delete(DeleteShiftDirection.xlShiftToLeft)


def del_cols(ibook, col_names):
    for sht in ibook.sheets:
        del_cols_in_one_sheet(sht, col_names)


def process_excel(ifile):
    app=xw.App(visible=True, add_book=False)
    app.display_alerts=False
    app.screen_updating=False
    print('Opening input file...')
    ibook = app.books.open(ifile)

    while True:
        col_names = input("Please input the col_names to delete (split by english ,), or input 'q' to exit:\n")
        if col_names.rstrip() == 'q':
            break
        col_names = col_names.rstrip().split(',')
        delete_cols(ibook)

    ibook.save()
    ibook.close()
    app.quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Extract some sheets from one excel and output selected sheets into a new excel')
    parser.add_argument('--ifile', action='store', required=True,  help='the input file')
    args = parser.parse_args()
    process_excel(args.ifile)
