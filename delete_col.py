#!/usr/bin/python3

import argparse
import xlwings as xw

def del_cols_in_one_sheet(sheet, col_names):
    print('===== Deleting columns in sheet: [{}] ====='.format(sheet.name))
    num_del = 0
    ncols = sheet[0, 0].expand("right").last_cell.column
    head_vals = sheet[0, 0:ncols].value
    if not head_vals:
        print('Empty header in sheet [{}], skip it'.format(sheet.name))
        return
    print('head_vals before deleting:\n{}'.format(head_vals))
    for col_name in col_names:
        try:
            col_id = head_vals.index(col_name)
        except ValueError:
            col_id = -1
        if col_id < 0:
            print('column [{}] not exist in sheet [{}], skip it'.format(col_name, sheet.name))
        else:
            col_letter = xw.utils.col_name(col_id + 1)  # xw.utils.col_name needs col_id starts from 1 but not 0
            sheet.range('{0}:{0}'.format(col_letter)).api.Delete()
            head_vals = sheet[0, 0:ncols].value
            num_del += 1
    ncols = sheet[0, 0].expand("right").last_cell.column
    head_vals = sheet[0, 0:ncols].value
    print('There are [{}] columns deleted. Head_vals after deleting:\n{}'.format(num_del, head_vals))


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
        col_names = input("Please input the col_names to delete (split by english ,), or input 'q' to save and exit:\n")
        if col_names.rstrip() == 'q':
            break
        col_names = col_names.rstrip().split(',')
        del_cols(ibook, col_names)

    print('Saving file...')
    ibook.save()
    ibook.close()
    app.quit()
    print('Done!')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Delete columns with supplied col_names from all sheets')
    parser.add_argument('--ifile', action='store', required=True,  help='the input file')
    args = parser.parse_args()
    process_excel(args.ifile)
