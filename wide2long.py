import sys, getopt
import xlwings as xw

def parse_excel(ifile, ofile, start_idx, step):
    app=xw.App(visible=False, add_book=False)
    app.display_alerts=False
    app.screen_updating=False

    ibook = app.books.open(ifile)
    isheet0 = ibook.sheets[0]
    last_cell = isheet0.used_range.last_cell
    nrows = last_cell.row
    ncols = last_cell.column
    print('sheet_name:%s, total_rows:%d, total_cols:%d' % (isheet0.name, nrows, ncols))

    input_data = isheet0[:nrows, :ncols].value
    output_data = []
    # Append the header row
    output_data.append(input_data[0][:start_idx+step])
    for in_row in range(1, nrows):
        common_info = input_data[in_row][:start_idx]
        for in_col in range(start_idx, ncols, step):
            # Skip empty cells inside one row. But if all cells are empty, we
            # still save the common info.
            if input_data[in_row][in_col] == None and in_col != start_idx:
                break
            output_data.append(common_info + input_data[in_row][in_col:in_col+step])

    obook = app.books.add()
    osheet0 = obook.sheets[0]
    osheet0.clear()
    osheet0[0, 0].options(expand='table').value = output_data
    obook.save(ofile)
    ibook.close()
    obook.close()
    app.quit()

def usage(myname):
    print('Usage: python3 %s --ifile <inputfile> --ofile <outputfile> --start <start_idx> --step <step_length>' % myname)

def main(argv):
    inputfile = ''
    outputfile = ''
    start_idx = 0
    step = 1
    try:
        opts, args = getopt.getopt(argv[1:],"i:o:h", ["ifile=","ofile=", "start=", "step="])
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
        elif opt in ("-o", "--ofile"):
            outputfile = arg
        elif opt == "--start":
            start_idx = int(arg)
        elif opt == "--step":
            step = int(arg)
    print ('input file is: ', inputfile)
    print ('output file is: ', outputfile)
    parse_excel(inputfile, outputfile, start_idx, step)

if __name__ == "__main__":
    main(sys.argv)
