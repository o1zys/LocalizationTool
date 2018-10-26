import os
import csv
import openpyxl.styles as sty
import openpyxl
import xlrd
import error


def copy_xls(old_xls):
    new_xls = openpyxl.Workbook()
    old_sheet = old_xls.sheet_by_index(0)
    new_sheet = new_xls.active
    max_row = old_sheet.nrows
    max_col = old_sheet.ncols

    for m in range(0, max_row):
        # for n in range(97, 97+max_col):
        for n in range(0, max_col):
            integer = n // 26
            remainder = n % 26
            if integer == 0:
                n_str = chr(remainder+97)
            else:
                n_str = chr(integer+96) + chr(remainder+97)

            i = '%s%d' % (n_str, m+1)
            cell = old_sheet.cell_value(m, n)
            new_sheet[i].value = cell
    return new_xls


def read_csv(path):
    arr = []
    if os.path.exists(path):
        file = csv.reader(open(path, 'r', encoding='UTF-8'))
        i = 0
        for line in file:
            arr.append([])
            for col in line:
                arr[i].append(col)
            i = i + 1
        return arr
    else:
        return


def combine_version(orig_list, version):
    if version == "":
        return orig_list
    parts = orig_list.split(',')
    add_flag = True
    for part in parts:
        if part.strip() == version:
            add_flag = False
            break
    if add_flag:
        if orig_list == "":
            return version
        else:
            return orig_list+', '+version
    return orig_list


def split_array_string(content, flag):
    str_list = []
    if not flag:
        str_list.append(content)
    else:
        parts = content.split(';')
        for part in parts:
            str_list.append(part)
    return str_list


# Convert to lua
def get_val_line(dict_sid, dict_row, i):
    for key, val in dict_row.items():
        if val == i:
            k = key
            break
    if dict_sid[k] == "":
        return int(dict_row[k])
    else:
        return int(dict_row[dict_sid[k]])
