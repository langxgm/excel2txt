#! /usr/bin/python3

import os
import sys
import getopt
import xlrd

def create_file(name, rows, header, is_print):
    # remove old file
    if os.path.exists(name):
        os.remove(name)
    # print header
    row_num = len(rows)
    column_num = len(rows[0])
    data_begin = header # data begin row
    fields = rows[0]
    types = rows[1]
    descs = rows[2]
    if is_print:
        print("create file: ", name)
        print("fields: ", fields)
        print("types : ", types)
        print("descs : ", descs)
        print("row_num: %d, column_num: %d, data_begin: %d" % (row_num, column_num, data_begin))
    # check header
    for n in range(column_num):
        field_name = fields[n]
        type_name = types[n]
        if field_name.isspace():
            print("create file: %s fail, field name is empty index=%d" % (name, n + 1))
            return
        if type_name.isspace():
            print("create file: %s fail, type name is empty index=%d field_name=%s" % (name, n + 1, field_name))
            return
        if type_name != "STRING" and type_name != "INT" and type_name != "FLOAT":
            print("create file: %s fail, type invalid index=%d field_name=%s type_name=%s" % (name, n + 1, field_name, type_name))
            return
    # out file
    file = open(name, "w", encoding="utf-16")
    try:
        for r in range(row_num):
            row_data = rows[r]
            if is_print:
                print(row_data)
            # "\n"
            if r > 0:
                file.write("\n")
            for c in range(column_num):
                # "\t"
                if c > 0:
                    file.write("\t")
                # field data
                field_name = fields[c]
                type_name = types[c]
                field_data = row_data[c]
                if r < data_begin:
                    file.write(str(field_data))
                else:
                    if type_name == "STRING":
                        assert (type(field_data) == str),\
                             "data is not a string type file: %s row=%d field_name=%s type_name=%s field_data=%s" % (name, r + 1, field_name, type_name, str(field_data))
                        file.write(field_data)
                    elif type_name == "INT":
                        assert (type(field_data) == float),\
                             "data is not a number type file: %s row=%d field_name=%s type_name=%s field_data=%s" % (name, r + 1, field_name, type_name, str(field_data))
                        assert (field_data - int(field_data) == 0),\
                             "data is not a int type file: %s row=%d field_name=%s type_name=%s field_data=%s" % (name, r + 1, field_name, type_name, str(field_data))
                        file.write(str(int(field_data)))
                    elif type_name == "FLOAT":
                        assert (type(field_data) == float),\
                             "data is not a float type file: %s row=%d field_name=%s type_name=%s field_data=%s" % (name, r + 1, field_name, type_name, str(field_data))
                        file.write(str(round(field_data, 4)))
    finally:
        # close file
        file.close()

class Usage(Exception):
    def __init__(self, msg):
        self.msg = msg

def help():
    print("""
    Help info
        -i, --input-dir         input file directory
        -o, --output-dir        output file directory
        -e, --filename-ext      filename extension, default is \".txt\"
        --header                table header, default is 3
    """)

def main(argv=None):
    if argv is None:
        argv = sys.argv
    count = 0
    try:
        current_dir = os.getcwd()
        input_dir = "."
        output_dir = "."
        filename_ext = ".txt"
        header = 3
        try:
            opts, args = getopt.getopt(argv[1:], "hi:o:e:", ["help", "input-dir=", "output-dir=", "filename-ext=", "header="])
            for opt_name, opt_value in opts:
                if opt_name in ("-h", "--help"):
                    help()
                    return 0
                elif opt_name in ("-i", "--input-dir"):
                    input_dir = opt_value
                elif opt_name in ("-o", "--output-dir"):
                    output_dir = opt_value
                elif opt_name in ("-e", "--filename-ext"):
                    filename_ext = opt_value
                elif opt_name == "--header":
                    header = int(opt_value)
        except getopt.error as msg:
            raise Usage(msg)

        for name in os.listdir(input_dir):
            # excel
            if name.endswith(".xlsx") or name.endswith(".xlsm"):
                table = xlrd.open_workbook(input_dir + "/" + name)
                #print("table: ", name)
                # sheet
                sheet_num = len(table.sheets())
                for sheet in table.sheets():
                    #print("sheet: ", sheet.name)
                    count += 1
                    # row count
                    nrows = sheet.nrows
                    # row data
                    rows = []
                    for i in range(nrows):
                        rows.append(sheet.row_values(i))
                    # check header
                    if len(rows) < header or (rows[1][0] != "INT" and rows[1][0] != "STRING"):
                        print("[%d] table: %s, sheet: %s, nrows: %d, ignore ! ! ! ! ! !" % (count, name, sheet.name, nrows))
                        continue
                    # create file
                    file_name = name.split(".", 1)[0]
                    if sheet.name.startswith("Sheet") == False:
                        file_name = sheet.name
                    create_file(output_dir + "/" + file_name + filename_ext, rows, header, False)
                    #print("[%d] table: %s, sheet: %s, nrows: %d, ok." % (count, name, sheet.name, nrows))
    except Usage as err:
        print(err.msg)
        help()
        os.system("pause")
        return 2
    except AssertionError as err:
       print("Assertion:", err)
       os.system("pause")
       return 2 
    except:
       print("Unexpected error:", sys.exc_info()[0])
       os.system("pause")
       return 2

    print("excel2txt: ", count)
    os.system("pause")
    return 0

if __name__ == "__main__":
    sys.exit(main())
