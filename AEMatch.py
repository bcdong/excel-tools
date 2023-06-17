import sys, getopt
import xlwings as xw
from datetime import datetime
from operator import itemgetter

kEventSheetName = "054 不良事件"
kExamSheetName = "001 实验室检查汇总"

kEventColNames = ["受试者代码", "不良事件名称", "CTCAE分级", "开始日期", "AE的转归", "转归日期"]
kExamColNames = ["受试者代码", "采样日期", "实验室检查项", "检测值", "参考值范围下限", "参考值范围上限", "临床意义值"]

def find_col_idx(col_name, col_name_list):
    for idx, col in enumerate(col_name_list):
        if col_name == col:
            return idx
    return -1

kEventUserIdIdx = find_col_idx("受试者代码", kEventColNames)
kEventNameIdx = find_col_idx("不良事件名称", kEventColNames)
kEventStartTimeIdx = find_col_idx("开始日期", kEventColNames)
kEventEndTimeIdx = find_col_idx("转归日期", kEventColNames)
kEventLevelIdx = find_col_idx("CTCAE分级", kEventColNames)
kEventResIdx = find_col_idx("AE的转归", kEventColNames)

kExamUserIdIdx = find_col_idx("受试者代码", kExamColNames)
kExamTimeIdx = find_col_idx("采样日期", kExamColNames)
kExamItemIdx = find_col_idx("实验室检查项", kExamColNames)
kExamValueIdx = find_col_idx("检测值", kExamColNames)
kExamLowBoundIdx = find_col_idx("参考值范围下限", kExamColNames)
kExamUpBoundIdx = find_col_idx("参考值范围上限", kExamColNames)
kExamMeaningIdx = find_col_idx("临床意义值", kExamColNames)

kMaxDatetime = datetime(2200, 1, 1, 0, 0, 0)

#  def ev_name2exam_item(ev_name):


# @param col_vals: The first element in each list of col_vals is the table header
# @ret: {userid: {event_name: [user_id, event_name, xxx, rowid]}}
def agg_by_2_dims(col_vals, dim1, dim2):
    nrows = len(col_vals[0])
    dict1 = {}
    for rowid in range(1, nrows): # iterate from 1 because line 0 is header
        item = []
        for col_val in col_vals:
            item.append(col_val[rowid])
        item.append("")  # error reason
        item.append(rowid + 1)  # add rowid to the last of a line item, +1 because excel start from 1 rather than 0
        item.append(True)   # this line is valid by default
        key1 = item[dim1]
        key2 = item[dim2]
        if key1 in dict1:
            dict2 = dict1[key1]
            if key2 in dict2:
                dict2[key2].append(item)
            else:
                dict2[key2]=[item]
        else:
            dict1[key1] = {key2: [item]}
    return dict1

def parse_datetime(time_str):
    try:
        time_date = datetime.strptime(time_str, '%Y-%m-%d')
        return (True, time_date)
    except ValueError as err:
        print("Error parsing date: %s. msg is: " % time_str, err)
        return (False, None)

def check_event(ev_list):
    for line in ev_list:
        start_time = line[kEventStartTimeIdx]
        if type(start_time) is str:
            if start_time[-2:] == "UN":
                start_time = start_time[:-2] + "01"
            time_tuple = parse_datetime(start_time)
            if time_tuple[0]:
                line[kEventStartTimeIdx] = time_tuple[1]
            else:
                # invalid start time, mark it invalid and stip this line
                line[kEventStartTimeIdx] = kMaxDatetime
                line[-1] = False
                line[-3] = "Invalid start time"
        end_time = line[kEventEndTimeIdx]
        if type(end_time) is str and end_time == "^":
            line[kEventEndTimeIdx] = kMaxDatetime
        elif type(end_time) is str:
            if end_time[-2:] == "UN":
                end_time = end_time[:-2] + "28"
            time_tuple = parse_datetime(end_time)
            if time_tuple[0]:
                line[kEventEndTimeIdx] = time_tuple[1]
            else:
                # invalid end time, mark it invalid and stip this line
                line[kEventEndTimeIdx] = kMaxDatetime
                line[-1] = False
                line[-3] = "Invalid end time"
    # sort by start time
    ev_list.sort(key=itemgetter(kEventStartTimeIdx))
    for idx, line in enumerate(ev_list):
        if line[-1] == False:  # this line is already invalid
            continue

        if line[kEventResIdx] == "未知":
            line[-1] = False
            line[-3] = "未知状态"
            continue

        if idx + 1 == len(ev_list):  # the last line
            if ((line[kEventResIdx] == "未恢复/未解决" and line[kEventEndTimeIdx] == kMaxDatetime) or
                    (line[kEventResIdx] == "恢复/解决" and line[kEventEndTimeIdx] != kMaxDatetime) or
                    (line[kEventResIdx] == "恢复/解决有后遗症" and line[kEventEndTimeIdx] != kMaxDatetime) or
                    (line[kEventResIdx] == "死亡" and line[kEventEndTimeIdx] != kMaxDatetime)):
                line[-1] = True
            else:
                line[-1] = False
                line[-3] = "最后一行状态不对"
            continue

        next_line = ev_list[idx + 1]
        if line[kEventResIdx] == "死亡":  # "死亡" must be the last line
            line[-1] = False
            line[-3] = "死亡状态不应该有下一行"
        elif line[kEventResIdx] == "恢复/解决" or line[kEventResIdx] == "恢复/解决有后遗症":
            if line[kEventEndTimeIdx] >= next_line[kEventStartTimeIdx]:
                line[-1] = False
                line[-3] = "恢复/解决结束时间晚于下一行开始时间"
        elif line[kEventResIdx] == "未恢复/未解决":
            if (line[kEventEndTimeIdx] != next_line[kEventStartTimeIdx]) or (line[kEventLevelIdx] >= next_line[kEventLevelIdx]):
                line[-1] = False
                line[-3] = "未恢复/未解决结束时间或分级与下一行不匹配"
        elif line[kEventResIdx] == "恢复中":
            if (line[kEventEndTimeIdx] != next_line[kEventStartTimeIdx]) or (line[kEventLevelIdx] <= next_line[kEventLevelIdx]):
                line[-1] = False
                line[-3] = "恢复中结束时间或分级与下一行不匹配"
        else:
            line[-1] = False
            line[-3] = "Unknown"

def mark_fail(line_list, err_msg):
    for line in line_list:
        line[-1] = -1
        line[-3] = err_msg

def check_one_user(ev_dict):
    for ev_name, ev_lists in ev_dict.items():
        if not ev_name:
            empty_lines = []
            for ev in ev_lists:
                empty_lines.append(ev[-2])
            print("Skip %d empty event, line numbers: " % len(ev_lists), empty_lines)
            continue
        check_event(ev_lists)

def check_event_user_dict(user_event_dict):
    for user_id, ev_dict in user_event_dict.items():
        check_one_user(ev_dict)

def read_one_sheet(isheet, col_names):
    out_cols = []
    valid_event_col_num = 0
    for i in range(len(col_names)):
        out_cols.append([])

    print("Starting to copy content of sheet:", isheet.name)
    nrows = isheet.used_range.last_cell.row
    ncols = isheet.used_range.last_cell.column
    headers = isheet[0, 0:ncols].value
    for idx, head_val in enumerate(headers):
        out_idx = find_col_idx(head_val, col_names)
        if out_idx == -1:
            continue
        col_vals = isheet[0:nrows, idx].value
        out_cols[out_idx] = col_vals
        valid_event_col_num += 1

    if valid_event_col_num != len(col_names):
        print("Invalid col_names: ", col_names)
        raise ValueError
    return out_cols


def read_excel(ifile, app):
    print("Opening excel file...")
    ibook = app.books.open(ifile)
    print("Excel file opened.")

    out_event_cols = []

    try:
        for isheet in ibook.sheets:
            if isheet.name == kEventSheetName:
                out_event_cols = read_one_sheet(isheet, kEventColNames)
    except ValueError:
        print("failt to copy excel, exiting...")
        ibook.close()
        app.quit()
        sys.exit(1)

    ibook.close()
    print("Excel file closed")
    return out_event_cols

def write_excel(ofile, app, data):
    obook = app.books.add()
    osheet0 = obook.sheets[0]
    osheet0.clear()
    osheet0[0, 0].options(expand='table').value = data
    obook.save(ofile)
    obook.close()

def usage(myname):
    print('Usage: python3 %s --ifile <inputfile> --ofile <outputfile>' % myname)

def main(argv):
    inputfile = ''
    try:
        opts, args = getopt.getopt(argv[1:],"i:o:h", ["ifile=", "ofile="])
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
    print ('input file is: ', inputfile)

    # Init excel app to read and write file
    app=xw.App(visible=False, add_book=False)
    app.display_alerts=False
    app.screen_updating=False

    event_col_vals = read_excel(inputfile, app)
    event_user_dict = agg_by_2_dims(event_col_vals, kEventUserIdIdx, kEventNameIdx)
    check_event_user_dict(event_user_dict)

    # reserve buffer when writing output to excel
    output_data = []
    # Append the header row
    outheader = kEventColNames[:2] + ["行号", "错误原因"] + kEventColNames[2:6]
    output_data.append(outheader)

    # print output to console
    print("======== Printing error records in %s:=========" % kEventSheetName)
    err_cnt = 0
    for user_id, ev_dict in event_user_dict.items():
        for ev_name, ev_lists in ev_dict.items():
            err_rowids = []
            for idx, ev in enumerate(ev_lists):
                if not ev[-1]:
                    err_rowids.append(idx)
                    err_cnt += 1
            if err_rowids:
                print("UserId: %s, EventName: %s" % (user_id, ev_name))
                for err_idx in err_rowids:
                    print("    rowid: %d, err: %s" %(ev_lists[err_idx][-2], ev_lists[err_idx][-3]))
                    out_line = [user_id, ev_name, ev_lists[err_idx][-2], ev_lists[err_idx][-3]]
                    out_line.extend(ev_lists[err_idx][2:6])
                    output_data.append(out_line)
    print("****** Total error count: %d ******" % err_cnt)

    write_excel(outputfile, app, output_data)
    app.quit()
    print("****** Writing output data to %s ******" % outputfile)

if __name__ == "__main__":
    main(sys.argv)
