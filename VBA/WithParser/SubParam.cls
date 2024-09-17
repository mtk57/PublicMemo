VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'メンバー変数
Private grep_results() As String

Private Const CLM_ = "B"
Private Const START_ROW_ = 13

'コンストラクタ
Private Sub Class_Initialize()
    Erase grep_results
End Sub

'初期化
Public Sub Init()
    CollectParam
End Sub

Private Sub CollectParam()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    Const LIMIT_EMPTY_ROWS = 5
    Dim row As Long: row = START_ROW_
    Dim i As Long: i = 0
    Dim empty_row_cnt As Long: empty_row_cnt = 0
    Dim is_before_empty_row As Boolean: is_before_empty_row = False
    Dim line As String
    
    
    Do
        line = main_sheet.Range(CLM_ & row + i).value
        
        If line = "" Then
            If is_before_empty_row = True Then
                empty_row_cnt = empty_row_cnt + 1
            Else
                empty_row_cnt = 1
                is_before_empty_row = True
            End If
        
            If empty_row_cnt >= LIMIT_EMPTY_ROWS Then
                Exit Do
            End If
            
            GoTo CONTINUE
        Else
            is_before_empty_row = False
        End If
        
        ReDim Preserve grep_results(i)
        grep_results(i) = line
        
CONTINUE:
        i = i + 1
    Loop
End Sub

Public Sub Validate()
    If GetGrepResultsCount() <= 0 Then
        Err.Raise 53, , "[SubParam] GREP結果が空です"
    End If
End Sub

'=========================================
'Getter
'=========================================
Public Function GetGrepResults() As String()
    GetGrepResults = grep_results
End Function

Public Function GetGrepResultsCount() As Long
    If Common.IsEmptyArrayLong(grep_results) = True Then
        GetGrepResultsCount = 0
        Exit Function
    End If
    
    Dim cnt As Long: cnt = UBound(grep_results)
    cnt = cnt + 1
    GetGrepResultsCount = cnt
End Function

