#! /usr/bin/python3

import os
import xlrd

current_dir = os.getcwd()
target_dir = "."
filename_ext = ".txt"

os.chdir(target_dir)
print(os.getcwd())

def create_file(name, rows, is_print):
    # remove old file
    if os.path.exists(name):
        os.remove(name)
    # print header
    row_num = len(rows)
    column_num = len(rows[0])
    data_begin = 3 # data begin row
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
            type_name = types[c]
            field_data = row_data[c]
            if r < data_begin:
                file.write(str(field_data))
            else:
                if type_name == "STRING":
                    assert(type(field_data) == type(""))
                    file.write(field_data)
                elif type_name == "INT":
                    assert(type(field_data) == type(0.0))
                    file.write(str(int(field_data)))
                elif  type_name == "FLOAT":
                    assert(type(field_data) == type(0.0))
                    file.write(str(round(field_data, 4)))
    # close file
    file.close()

count = 0
for name in os.listdir("."):
    # excel
    if name.endswith(".xlsx") or name.endswith(".xlsm"):
        table = xlrd.open_workbook(name)
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
            if len(rows) < 3 or (rows[1][0] != "INT" and rows[1][0] != "STRING"):
                print("[%d] table: %s, sheet: %s, nrows: %d, ignore ! ! ! ! ! !" % (count, name, sheet.name, nrows))
                continue
            # create file
            file_name = name.split(".", 1)[0]
            if sheet.name.startswith("Sheet") == False:
                file_name = sheet.name
            create_file(file_name + filename_ext, rows, False)
            #print("[%d] table: %s, sheet: %s, nrows: %d, ok." % (count, name, sheet.name, nrows))

print("excel2txt: ", count)

os.system("pause")