VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�����o�[�ϐ�

Private target_sheet_name As String
Private grep_result_cell As String
Private find_word_cell As String
Private find_condition_cell As String
Private replace_char As String
Private is_regex_str As String
Private is_regex As Boolean

Private grep_result_cell_clm As String
Private grep_result_cell_row As Long
Private find_word_cell_clm As String
Private find_word_cell_row As Long
Private find_condition_cell_clm As String
Private find_condition_cell_row As Long

Private Const CLM_ = "O"
Private Const ROW_ = 6

'�R���X�g���N�^
Private Sub Class_Initialize()
    target_sheet_name = ""
    grep_result_cell = ""
    find_word_cell = ""
    find_condition_cell = ""
    replace_char = ""
    is_regex = False
End Sub

'������
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    target_sheet_name = main_sheet.Range(CLM_ & ROW_ + 0).value
    grep_result_cell = main_sheet.Range(CLM_ & ROW_ + 1).value
    find_word_cell = main_sheet.Range(CLM_ & ROW_ + 2).value
    find_condition_cell = main_sheet.Range(CLM_ & ROW_ + 3).value
    replace_char = main_sheet.Range(CLM_ & ROW_ + 4).value
    is_regex_str = main_sheet.Range(CLM_ & ROW_ + 5).value
End Sub

'����
Public Sub Validate()
    If target_sheet_name = "" Then
        Err.Raise 53, , "�ΏۃV�[�g�����w�肳��Ă��܂���"
    ElseIf Common.IsExistSheet(ThisWorkbook, target_sheet_name) = False Then
        Err.Raise 53, , "�ΏۃV�[�g�����݂��܂���(" & target_sheet_name & ")"
    End If
    
    If grep_result_cell = "" Then
        Err.Raise 53, , "Grep���ʂ̊J�n�Z�����w�肳��Ă��܂���"
    End If
    
    If Common.SplitCellAddress(grep_result_cell, grep_result_cell_clm, grep_result_cell_row) = False Then
        Err.Raise 53, , "Grep���ʂ̊J�n�Z�����s���ł�(" & grep_result_cell & ")"
    End If
    
    If find_word_cell = "" Then
        Err.Raise 53, , "�������[�h�̊J�n�Z�����w�肳��Ă��܂���"
    End If
    
    If Common.SplitCellAddress(find_word_cell, find_word_cell_clm, find_word_cell_row) = False Then
        Err.Raise 53, , "�������[�h�̊J�n�Z�����s���ł�(" & find_word_cell & ")"
    End If
    
    If find_condition_cell = "" Then
        Err.Raise 53, , "���������̊J�n�Z�����w�肳��Ă��܂���"
    End If
    
    If Common.SplitCellAddress(find_condition_cell, find_condition_cell_clm, find_condition_cell_row) = False Then
        Err.Raise 53, , "���������̊J�n�Z�����s���ł�(" & find_condition_cell & ")"
    End If

    If replace_char = "" Then
        Err.Raise 53, , "���������̒u���O�������w�肳��Ă��܂���"
    End If
    
    If is_regex_str = "" Or is_regex_str = "NO" Then
        is_regex = False
    Else
        is_regex = True
    End If
End Sub

'=========================================
'Getter
'=========================================

Public Function GetTargetSheetName() As String
    GetTargetSheetName = target_sheet_name
End Function

Public Function GetGrepResultCell() As String
    GetGrepResultCell = grep_result_cell
End Function

Public Function GetGrepResultCellColumn() As String
    GetGrepResultCellColumn = grep_result_cell_clm
End Function

Public Function GetGrepResultCellRow() As Long
    GetGrepResultCellRow = grep_result_cell_row
End Function

Public Function GetFindWordCell() As String
    GetFindWordCell = find_word_cell
End Function

Public Function GetFindWordCellColumn() As String
    GetFindWordCellColumn = find_word_cell_clm
End Function

Public Function GetFindWordCellRow() As Long
    GetFindWordCellColumn = find_word_cell_row
End Function

Public Function GetFindConditionCell() As String
    GetFindConditionCell = find_condition_cell
End Function

Public Function GetFindConditionCellColumn() As String
    GetFindConditionCellColumn = find_condition_cell_clm
End Function

Public Function GetFindConditionCellRow() As Long
    GetFindConditionCellRow = find_condition_cell_row
End Function

Public Function GetReplaceChar() As String
    GetReplaceChar = replace_char
End Function

Public Function IsRegEx() As Boolean
    IsRegEx = is_regex
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "target_sheet_name=(" & target_sheet_name & "), " & _
          "grep_result_cell=(" & grep_result_cell & "), " & _
          "find_word_cell=(" & find_word_cell & "), " & _
          "find_condition_cell=(" & find_condition_cell & "), " & _
          "replace_char=(" & replace_char & "), " & _
          "is_regex=(" & is_regex_str & ")"

    GetAllValue = ret
End Function