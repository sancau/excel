import os, sys

"""
Data config
"""
INDEX = 0
NAME = 1
SIZE = 6
AMOUNT = 11
MATERIAL = 12
UNITS = 17
MATERIAL_AMOUNT = 19
STANDART = 21
COMMENT = 24

PAYLOAD_DATA_INDEXES = [
    INDEX, NAME, SIZE, 
    AMOUNT, MATERIAL, UNITS, 
    MATERIAL_AMOUNT, STANDART, 
    COMMENT
]

FISRT_LIST = 'Лист1'
FISRT_LIST_FIRST_DATA_ROW = 22
OTHER_LISTS_FIRST_DATA_ROW = 2

REPEAT_SYMBOLS = ['——ıı——', ]

AMOUNT_UNITS_VERBOSE = 'шт'

SIZE_DICT = {
    '≠': {
        'verbose': 'ТОЛЩ.',
        'units': 'ММ',
    },
    'Ø': {
        'verbose': 'ДИАМ.',
        'units': 'ММ'
    }
}

UNDEFINED_SYMBOL = '(нет данных)'

"""
Script config
"""
DEFAULT_DIR_PATH = os.path.dirname(sys.argv[0])
DEFAULT_RESULT_FILE_NAME = 'output.xlsx'
DEFAULT_RESULT_FILE_PATH = os.path.join(DEFAULT_DIR_PATH, DEFAULT_RESULT_FILE_NAME)

