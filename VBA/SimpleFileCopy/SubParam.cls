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
Private file_paths() As String

Private Const CLM_ = "C"
Private Const START_ROW_ = 14

'コンストラクタ
Private Sub Class_Initialize()
    Erase file_paths
End Sub

'初期化
Public Sub Init()
    CollectParam
End Sub

Private Sub CollectParam()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    Dim row As Long: row = START_ROW_
    Dim i As Long: i = 0
    Dim path As String
    
    Do
        path = main_sheet.Range(CLM_ & row + i).value
        
        If path = "" Then
            Exit Do
        End If
        
        ReDim Preserve file_paths(i)
        file_paths(i) = path
        
        i = i + 1
    Loop
End Sub

'検証
Public Sub Validate()
    'Do nothing
End Sub

'=========================================
'Getter
'=========================================

Public Function GetFilePathList() As String()
    GetFilePathList = file_paths()
End Function

Public Function GetFilePathListCount() As Long
    If Common.IsEmptyArray(file_paths) = True Then
        GetFilePathListCount = 0
        Exit Function
    End If
    
    Dim cnt As Long: cnt = UBound(file_paths())
    cnt = cnt + 1
    GetFilePathListCount = cnt
End Function

