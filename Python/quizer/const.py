from enum import IntEnum

VERSION = r'ver=1.0'
VER_POS = 'A1'

# クイズファイルパス
# DEFAULT_EXCEL_FILE_NAME = r'sample.xlsx'
DEFAULT_EXCEL_FILE_NAME = r'C:\_git\Memo\Other\Salesforce\QUIZ.xlsx'

# 統計情報DBファイルパス
# SQLITE3
DEFAULT_DB_FILE_NAME = r'C:\_git\Memo\Other\Salesforce\QUIZ.db'

DEFAULT_QUESTION_NUM = 60   # 問題数
DEFAULT_PASS_LINE = 65      # 合格ライン(%)

SHEET_COMMON = 'common'
SHEET_QUIZ_ADMIN = 'アドミニストレーター'
OFFSET_ADMIN = 4

REQUIRED_SHEETS = [SHEET_COMMON, SHEET_QUIZ_ADMIN]

MARK_COMMA = ','
MARK_SP = ' '
MARK_PLUS = '+'
SPLITS = r'[ \+]'

MARK_CORRECT = '○'  # 正解マーク

MARK_Y = 'Y'
MARK_N = 'N'

MIN_ANSWER = 1      # 回答できる最小数
MAX_ANSWER = 8      # 回答できる最大数

MAX_COLUMNS = 4


class Offset(IntEnum):
    """ シートの列名の桁位置オフセット """
    NUM = 1         # 項番
    IS_SKIP = 2     # SKIP?
    QUESTION = 3    # 問題
    CHOICE = 4      # 選択肢
    ANSWER = 5      # 正解


class Mode(IntEnum):
    """ モード """
    QUIZ = 1        # クイズモード
    LEARN = 2       # 学習モード
