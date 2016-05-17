
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

def merge(rows):
    """
    Parse a row of the table.
    Add data to result dict.
    
    If this key exists in the dict process its params
    and add the amount
    Else - add new key in dict
    """
    # clean data from openpyxl wrappers
    plain_rows = []
    for row in rows:
        plain_row = [item.value for item in row]               
        plain_rows.append(plain_row)
    
    #handle repeat sybmols
    for row_index, plain_row in enumerate(plain_rows):
        for value_index, value in enumerate(plain_row):
            if value in REPEAT_SYMBOLS:
                plain_row[value_index] = plain_rows[row_index - 1][value_index] 
    
    filtered_rows = [
        row for row in plain_rows if row[NAME] is not None and row[INDEX] is not None
    ] 
       
    def remove_spaces(value):
        """
        Returns given string with all the spaces removed.
        """
        return re.sub('[\s+]', '', str(value))
    
    def get_size(row):
        """
        Returns main size criteria for a given row.
        """
        return remove_spaces(str(row[SIZE])).split('×')[0]
           
    def get_match(row, output):
        """
        Checks if given row has a match in output array.
        """
        standart_match = [r for r in output if remove_spaces(r[STANDART]) == remove_spaces(row[STANDART])]
        if not standart_match:
            output.append(row)
            return None
        else:
            material_match = [r for r in standart_match if remove_spaces(r[MATERIAL]) == remove_spaces(row[MATERIAL])]
            if not material_match:
                output.append(row)
                return None
            else:
                size_match = [r for r in material_match if get_size(r) == get_size(row)]
                if not size_match:
                    output.append(row)
                    return None
                else:
                    print('Found %s merge candidates!' % len(size_match))
                    if len(size_match) > 1:
                        raise Exception('MORE THEN 1 MERGE CANDIDATE WAS FOUND. PARSING ALGORYTHM IS INVALID!') 
                    return size_match[0]
                 
    def merge_data(new, existed):
        """
        Appends row data to an existed row if match was confirmed.
        
        Apply merge politics here to inject new in existed.
        """
        print('Merging...')
        if existed[NAME] != new[NAME]:
            existed[NAME] = ', '.join([existed[NAME], new[NAME]])
        size_amount_existed = ', '.join([str(existed[SIZE]), str(existed[AMOUNT])])
        size_amount_new = ', '.join([str(new[SIZE]), str(new[AMOUNT])])
        existed[SIZE] = '; '.join([size_amount_existed, size_amount_new])
        return existed
  
    output = []
    print('Processing output...')
    for row in filtered_rows:
        match = get_match(row, output)
        print(len(output))
        if match:
            output[output.index(match)] = merge_data(row, match)
    
    ready_output = []
    index = 1
    for row in [[cell for cell_index, cell in enumerate(row) if cell_index in PAYLOAD_DATA_INDEXES] for row in output]:
        row[0] = index
        ready_output.append(row)
        index += 1
    return ready_output

def build_results_file(rows, result_file_path):
    """
    Build an excel file based on results dict and 
    a given path.
    """
    wb = openpyxl.load_workbook('template.xlsx')
    dest_filename = os.path.join(result_file_path, DEFAULT_RESULT_FILE_NAME)
    ws = wb.active   
    for row in rows:
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
                            rows_to_process.append(row) 
                            
            result = merge(rows_to_process)
            build_results_file(result, result_file_path) 
            print('Success')
        else:
            print('No files to process')
    
    except Exception as ex:
        print('Error while processing')
        print(ex)
