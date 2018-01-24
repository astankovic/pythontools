from os import path
from os import listdir
from openpyxl import load_workbook
import re

def read_the_folder_for_xlsx():
    xlsx_files = list()
    file_list = listdir(path.curdir)
    for i in file_list:
        if path.isfile(i) & (path.splitext(i)[1] == ".xlsx"):
            xlsx_files.append(path.abspath(i))
    return xlsx_files

def vlookup_across_multiple_xls(term, search_column, rel_lookup_column, include_file_name_in_output):
    for r in read_the_folder_for_xlsx():
        wb = load_workbook(r)
        for sheet in wb:
            for i in range(1,sheet.max_row + 1):
                iterator = sheet.cell(row = i, column = search_column)
                if iterator.value == term:
                    print('{}\t{}\t{}'.format(sheet.cell(row = i, column = search_column + rel_lookup_column).value, sheet.title, path.basename(r)))

                        
searchText = input("Search term: ")
vlookup_across_multiple_xls(searchText, 2, 2, False)
text = input("Good with this?")