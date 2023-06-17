import sys, getopt
import xlwings as xw

def do_filter(sheets, col_name, val):
    for isheet in sheets:
        ncols = isheet.used_range.last_cell.column
        head_vals = isheet[0, 0:ncols].value
        col_id = -1
        for i in range(ncols):
            if head_vals[i] == col_name:
                col_id = i
                break
        if col_id >= 0:
            isheet.used_range.api.AutoFilter(Field:=col_id+1, Criterial:=val)
            print("Filter success on sheet: %s" % isheet.name)

def do_reset(sheets):
    for isheet in sheets:
        isheet.api.AutoFilterMode = False

def filter_excel(ifile):
    app=xw.App(visible=True, add_book=False)
    app.display_alerts=True
    app.screen_updating=True
    ibook = app.books.open(ifile)
    sheets = ibook.sheets

    while True:
        cmd_str = input("Please input the col_name and expected_value to filter, or input 'q' to exit, or input 'r' to reset filter:\n")
        if cmd_str.rstrip() == 'q':
            break
        if cmd_str.rstrip() == 'r':
            do_reset(sheets)
            continue
        cmd_list = cmd_str.rstrip().split()
        if len(cmd_list) != 2:
            print("Please input the col_name and expected_value seperated by a space")
            continue
        do_filter(sheets, cmd_list[0], cmd_list[1])

    ibook.close()
    app.quit()

def usage(myname):
    print('Usage: python3 %s --ifile <inputfile>' % myname)

def main(argv):
    inputfile = ''
    try:
        opts, args = getopt.getopt(argv[1:],"i:h", ["ifile="])
    except getopt.GetoptError as err:
        print('Error: ', err)
        usage(argv[0])
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            usage(argv[0])
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
    print ('input file is: ', inputfile)
    filter_excel(inputfile)

if __name__ == "__main__":
    main(sys.argv)
