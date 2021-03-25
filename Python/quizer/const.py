from enum import IntEnum

VERSION = r'ver=0.6'
VER_POS = 'A1'

# DEFAULT_EXCEL_FILE_NAME = r'sample.xlsx'
DEFAULT_EXCEL_FILE_NAME = r'C:\_git\Memo\Other\Salesforce\QUIZ.xlsx'
DEFAULT_QUESTION_NUM = 60   # 問題数
DEFAULT_PASS_LINE = 65      # 合格ライン(%)

SHEET_COMMON = 'common'
SHEET_QUIZ_ADMIN = 'アドミニストレーター'
OFFSET_ADMIN = 4

REQUIRED_SHEETS = [SHEET_COMMON, SHEET_QUIZ_ADMIN]

MARK_COMMA = ','
MARK_SP = ' '
MARK_PLUS = '+'
SPLITS = [MARK_SP, MARK_PLUS]

MARK_CORRECT = '○'  # 正解マーク

MIN_ANSWER = 1      # 回答できる最小数
MAX_ANSWER = 8      # 回答できる最大数

MAX_COLUMNS = 4


class Offset(IntEnum):
    NUM = 1
    QUESTION = 2
    CHOICE = 3
    ANSWER = 4


class Mode(IntEnum):
    QUIZ = 1
    LEARN = 2