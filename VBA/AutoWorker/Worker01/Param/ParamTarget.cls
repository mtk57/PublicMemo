VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private HEADER As String

'メンバー変数
Private vbprj_file_path As String
Private lot As String
Private phase As String
Private pid As String
Private ticket_num As String
Private branch As String
Private commit As String
Private tag As String

'コンストラクタ
Private Sub Class_Initialize()
    HEADER = "[Target Param] "
    
    vbprj_file_path = ""
    lot = ""
    phase = ""
    pid = ""
    ticket_num = ""
    branch = ""
    commit = ""
    tag = ""
End Sub

'初期化
Public Sub Init(ByVal row As Integer)
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("params")
    
    Dim C1 As String: C1 = Define.TARGET_PARAM_CLM_1
    Dim C2 As String: C2 = Define.TARGET_PARAM_CLM_2
    Dim C3 As String: C3 = Define.TARGET_PARAM_CLM_3
    Dim C4 As String: C4 = Define.TARGET_PARAM_CLM_4
    Dim C5 As String: C5 = Define.TARGET_PARAM_CLM_5
    Dim C6 As String: C6 = Define.TARGET_PARAM_CLM_6
    Dim C7 As String: C7 = Define.TARGET_PARAM_CLM_7
    Dim C8 As String: C8 = Define.TARGET_PARAM_CLM_8
    
    vbprj_file_path = sheet.Range(C1 & row).value
    lot = sheet.Range(C2 & row).value
    phase = sheet.Range(C3 & row).value
    pid = sheet.Range(C4 & row).value
    ticket_num = sheet.Range(C5 & row).value
    branch = sheet.Range(C6 & row).value
    commit = sheet.Range(C7 & row).value
    tag = sheet.Range(C8 & row).value
End Sub

'検証
Public Sub Validate(ByVal row As Integer)
    If vbprj_file_path = "" Then
        Err.Raise 53, , HEADER & "VBプロジェクトファイルが指定されていません(row=" & row & ")"
    End If
    
    If Common.IsExistsFile(vbprj_file_path) = False Then
        Err.Raise 53, , HEADER & "VBプロジェクトファイルが見つかりません(vbprj_file_path=" & vbprj_file_path & ", row=" & row & ")"
    End If
    
    If lot = "" Then
        Err.Raise 53, , HEADER & "Lotが指定されていません(row=" & row & ")"
    End If
    
    If phase = "" Then
        Err.Raise 53, , HEADER & "Phaseが指定されていません(row=" & row & ")"
    End If
    
    If pid = "" Then
        Err.Raise 53, , HEADER & "PIDが指定されていません(row=" & row & ")"
    End If
    
    If ticket_num = "" Then
        Err.Raise 53, , HEADER & "ticket_numが指定されていません(row=" & row & ")"
    End If
    
    If branch = "" Then
        Err.Raise 53, , HEADER & "branchが指定されていません(row=" & row & ")"
    End If
    
    If commit = "" Then
        Err.Raise 53, , HEADER & "commitが指定されていません(row=" & row & ")"
    End If

    If tag = "" Then
        Err.Raise 53, , HEADER & "tagが指定されていません(row=" & row & ")"
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

Public Function GetLot() As String
    GetLot = lot
End Function

Public Function GetPhase() As String
    GetPhase = phase
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
          "lot=(" & lot & "), " & _
          "phase=(" & phase & "), " & _
          "pid=(" & pid & "), " & _
          "ticket_num=(" & ticket_num & "), " & _
          "branch=(" & branch & "), " & _
          "commit=(" & commit & "), " & _
          "tag=(" & tag & ")"

    GetAllValue = ret
End Function



