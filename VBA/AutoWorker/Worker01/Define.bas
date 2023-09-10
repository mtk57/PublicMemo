Attribute VB_Name = "Define"
Option Explicit

Public Const DEBUG_LOG_CELL = "D8"
Public Const NOW_PROCESS = "A2"

Public Const COMMON_PARAM_CLM = "D"
Public Const COMMON_PARAM_ROW = 5
Public Const COMMON_PARAM_ROWS = 6

Public Const GIT_PARAM_CLM = "D"
Public Const GIT_PARAM_ROW = 13
Public Const GIT_PARAM_ROWS = 7

Public Const TARGET_PARAM_CLM_0 = "B"
Public Const TARGET_PARAM_CLM_1 = "C"
Public Const TARGET_PARAM_CLM_2 = "D"
Public Const TARGET_PARAM_CLM_3 = "E"
Public Const TARGET_PARAM_CLM_4 = "F"
Public Const TARGET_PARAM_CLM_5 = "G"
Public Const TARGET_PARAM_CLM_6 = "H"
Public Const TARGET_PARAM_CLM_7 = "I"
Public Const TARGET_PARAM_CLM_8 = "J"
Public Const TARGET_PARAM_CLM_9 = "K"
Public Const TARGET_PARAM_ROW = 24

Public Enum PROCESS_TYPE
    UNKNOWN
    PROC_001
    PROC_002
    PROC_003
    PROC_004
    PROC_005
    DELETE_BRANCH
    DELETE_TAG
End Enum

