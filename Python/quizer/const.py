from enum import IntEnum

VERSION = r'ver=0.1'

VER_POS = 'A1'

# DEFAULT_EXCEL_FILE_NAME = r'sample.xlsx'
DEFAULT_EXCEL_FILE_NAME = r'C:\_git\Memo\Other\Salesforce\QUIZ_administrator.xlsx'

SHEET_QUIZ = 'QUIZ'
DEFAULT_SHEET_NAME = SHEET_QUIZ

REQUIRED_SHEETS = [SHEET_QUIZ]


GROUP_ADMIN = 'QUIZ'
OFFSET_ADMIN = 4


class Offset(IntEnum):
    NUM = 1
    QUESTION = 2
    CHOICE = 3
    ANSWER = 4


MARK_COMMA = ','
MARK_CORRECT = '○'  # 正解マーク

MIN_ANSWER = 1      # 回答できる最小数
MAX_ANSWER = 8      # 回答できる最大数
