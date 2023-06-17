#!/usr/bin/python3

import argparse
import xlwings as xw

def copy_sheets(app, ibook, ofile):
    all_sheets = ibook.sheet_names
    print('All existing sheets: {}'.format(all_sheets))
    out_sheets = input("Please input the sheet names (split by spaces) to extract, or input 'q' to exit:\n")
    if out_sheets.rstrip() == 'q':
        return
    out_sheets = out_sheets.rstrip().split()
    obook = app.books.add()
    out_idx = 0
    for sheet_name in out_sheets:
        if sheet_name not in all_sheets:
            print('Invalid sheet name: {}. Continue to copy next sheet.'.format(sheet_name))
            continue
        ibook.sheets[sheet_name].copy(after=obook.sheets[out_idx])
        out_idx += 1
    obook.save(ofile)
    obook.close()
    print('Copy {} sheets done!'.format(out_idx))


def process_excel(ifile, ofile):
    app=xw.App(visible=False, add_book=False)
    app.display_alerts=False
    app.screen_updating=False
    ibook = app.books.open(ifile)
    copy_sheets(app, ibook, ofile)
    ibook.close()
    app.quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Extract some sheets from one excel and output selected sheets into a new excel')
    parser.add_argument('--ifile', action='store', required=True,  help='the input file')
    parser.add_argument('--ofile', action='store', required=True,  help='the output file')
    args = parser.parse_args()

    process_excel(args.ifile, args.ofile)
