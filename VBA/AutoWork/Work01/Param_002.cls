VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Param_002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private HEADER As String

'メンバー変数
Private vbprj_file_path As String
Private pid As String
Private ticket_num As String
Private branch As String
Private commit As String
Private tag As String

'コンストラクタ
Private Sub Class_Initialize()
    HEADER = "[Param002] "
    
    vbprj_file_path = ""
    pid = ""
    ticket_num = ""
    branch = ""
    commit = ""
    tag = ""
End Sub

'初期化
Public Sub Init(ByVal ROW As Integer)
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(Define.SHEET_01)
    
    vbprj_file_path = main_sheet.Range("C" & ROW).value
    pid = main_sheet.Range("D" & ROW).value
    ticket_num = main_sheet.Range("E" & ROW).value
    branch = main_sheet.Range("F" & ROW).value
    commit = main_sheet.Range("G" & ROW).value
    tag = main_sheet.Range("H" & ROW).value
End Sub

'検証
Public Sub Validate(ByVal ROW As Integer)
    If vbprj_file_path = "" Then
        Err.Raise 53, , HEADER & "VBプロジェクトファイルが指定されていません(row=" & ROW & ")"
    End If
    
    If Common.IsExistsFile(vbprj_file_path) = False Then
        Err.Raise 53, , HEADER & "VBプロジェクトファイルが見つかりません(vbprj_file_path=" & vbprj_file_path & ", row=" & ROW & ")"
    End If
    
    If pid = "" Then
        Err.Raise 53, , HEADER & "pidが指定されていません(row=" & ROW & ")"
    End If
    
    If ticket_num = "" Then
        Err.Raise 53, , HEADER & "ticket_numが指定されていません(row=" & ROW & ")"
    End If
    
    If branch = "" Then
        Err.Raise 53, , HEADER & "branchが指定されていません(row=" & ROW & ")"
    End If
    
    If commit = "" Then
        Err.Raise 53, , HEADER & "commitが指定されていません(row=" & ROW & ")"
    End If

    If tag = "" Then
        Err.Raise 53, , HEADER & "tagが指定されていません(row=" & ROW & ")"
    End If
End Sub

'=========================================
'Getter
'=========================================

Public Function GetVBPrjFilePath() As String
    GetVBPrjFilePath = vbprj_file_path
End Function

Public Function GetTicketNum() As String
    GetTicketNum = ticket_num
End Function

Public Function GetPID() As String
    GetPID = pid
End Function

Public Function GetBranch() As String
    GetBranch = branch
End Function

Public Function GetCommit() As String
    GetCommit = commit
End Function

Public Function GetTag() As String
    GetTag = tag
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = HEADER & _
          "vbprj_file_path=(" & vbprj_file_path & "), " & _
          "ticket_num=(" & ticket_num & "), " & _
          "pid=(" & pid & "), " & _
          "branch=(" & branch & "), " & _
          "commit=(" & commit & "), " & _
          "tag=(" & tag & ")"

    GetAllValue = ret
End Function



