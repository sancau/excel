
import re
import os, sys
import openpyxl

from openpyxl.compat import range
from openpyxl.cell import get_column_letter

from config import *

"""
LOGIC
"""    
def get_files(dir_path=DEFAULT_DIR_PATH):
    """
    Returns list of excel file paths in
    a given directory
    """
    result = []
    for path, subdirs, files in os.walk(dir_path):
        for file in files:
            filename, file_extension = os.path.splitext(file)
            if file_extension in ['.xls', '.xlsx']:
                file_path = os.path.join(path, file)
                result.append(file_path)   
    return result

def getValueWithMergeLookup(sheet, cell):
    idx = cell.coordinate
    for range_ in sheet.merged_cell_ranges:
        merged_cells = list(openpyxl.utils.rows_from_range(range_))
        for row in merged_cells:
            if idx in row:
                # If this is a merged cell,
                # return  the first cell of the merge range
                return sheet.cell(merged_cells[0][0]).value

    return sheet.cell(idx).value

def merge(rows):
    """
    Parse a row of the table.
    Add data to result dict.
    
    If this key exists in the dict process its params
    and add the amount
    Else - add new key in dict
    """      
    filtered_rows = [
        row for row in rows if row[NAME] is not None and row[INDEX] is not None
    ] 
    
    payloaded = []
    for row in filtered_rows:
        payloaded_row = []
        if type(row[0]) == int:
            for idx in PAYLOAD_DATA_INDEXES:
                payloaded_row.append(row[idx])
            payloaded.append(payloaded_row)
               
    #handle repeat sybmols
    for row_index, row in enumerate(payloaded):
        for value_index, value in enumerate(row):
            if value in REPEAT_SYMBOLS:
                row[value_index] = payloaded[row_index - 1][value_index] 
            elif value is None:
                row[value_index] = '-'

    def remove_spaces(value):
        """
        Returns given string with all the spaces removed.
        """
        return re.sub('[\s+]', '', str(value))
    
    def get_primary_size(size_str):
        """
        Returns main size criteria for a given row.
        """
        return remove_spaces(str(row[SIZE])).split('×')[0]
           
    def merge(output, row):
        """
        Checks if given row has a match in output array.
        """        
        name_idex = 0
        size_index = 2
        standart_index = 6
        units_index = 7
        amount_index = 8
        material_index = 10

        def _(value):
            if value:
                return str(value)
            return '-'
        for item in row:
            row[row.index(item)] = _(item)

        def get_standard_match(array, candidate):
            if not array:
                return []
            return [i for i in array if remove_spaces(i[standart_index]) \
                == remove_spaces(candidate[STANDART])]

        def get_material_match(array, candidate):
            return [i for i in array if remove_spaces(i[material_index]) \
                == remove_spaces(candidate[MATERIAL])]

        def get_primary_size_match(array, candidate):
            return 0

        def size_equal(array_row, candidate):
            return False

        def make_new(candidate):
            new = [
                row[NAME] + ', ' + row[MATERIAL],
                '-',
                [row[SIZE] + ', ' + row[AMOUNT] + ' ' + row[UNITS] + ';', ],
                '-',
                '-',
                '-',
                row[STANDART],
                row[UNITS],
                row[AMOUNT],
                '-',
                row[MATERIAL]
            ]
            output.append(new)

        standart_match = get_standard_match(output, row)
        if standart_match:
            print('совпадений по ГОСТ: %s' % len(standart_match))
            material_match = get_material_match(standart_match, row)
            if material_match:
                print('совпадений по МАТЕРИАЛ: %s' % len(material_match))
                primary_size_match = get_primary_size_match(material_match, row)
                if primary_size_match:
                    print('FOUND MERGE CANDIDATE')
                    if size_equal(primary_size_match, row):
                        # SUM AMOUNTS
                        pass
                    else:
                        # ADD SIZE / AMOUNT
                        pass
        make_new(row)
        return output
                  
    output = []
    print('Processing output...')
    
    # TODO 

    # for row in payloaded:
    #     output = merge(output, row)
    return payloaded

def build_results_file(rows, result_file_path):
    """
    Build an excel file based on results dict and 
    a given path.
    """
    wb = openpyxl.load_workbook('template.xlsx')
    dest_filename = os.path.join(result_file_path, DEFAULT_RESULT_FILE_NAME)
    ws = wb.active   
    for row in rows:
        for value_index, value in enumerate(row):
            row[value_index] = str(value).encode('utf-8')
        ws.append(row)   
    wb.save(filename = dest_filename)
   
def process_files(dir_path=DEFAULT_DIR_PATH, result_file_path=DEFAULT_RESULT_FILE_PATH):
    try:
        files = get_files(dir_path)
        rows_to_process = []
        if files:
            for file in files:
                workbook = openpyxl.load_workbook(filename=file)   
                for sheet in workbook:
                    for row in sheet:
                        # add rows by certain condition
                        row_index = (lambda x: x[0].row)(row)
                        if (sheet is not workbook[FISRT_LIST] \
                         and row_index >= OTHER_LISTS_FIRST_DATA_ROW) \
                            or row_index >= FISRT_LIST_FIRST_DATA_ROW:
                            
                            merged_cells_awared_row = []
                            for cell in row:
                                value = getValueWithMergeLookup(sheet, cell)
                                merged_cells_awared_row.append(value)
                                                       
                            rows_to_process.append(merged_cells_awared_row) 
            
            result = merge(rows_to_process)
            build_results_file(result, result_file_path) 
            print('Success')
        else:
            print('No files to process')
    
    except Exception as ex:
        print('Error while processing')
        print(ex)