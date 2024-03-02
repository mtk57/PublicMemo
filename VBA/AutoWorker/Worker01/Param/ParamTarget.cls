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
Private proc_type As PROCESS_TYPE

Private is_enable As Boolean
Private is_enable_s As String
Private vbprj_file_path As String
Private lot As String
Private phase As String
Private pid As String
Private ticket_num As String
Private step As String
Private dest_base_dir As String
Private branch As String
Private commit As String
Private tag As String
Private any1 As String


'コンストラクタ
Private Sub Class_Initialize()
    HEADER = "[Target Param] "
    
    proc_type = UNKNOWN
    is_enable = False
    vbprj_file_path = ""
    lot = ""
    phase = ""
    pid = ""
    ticket_num = ""
    step = ""
    dest_base_dir = ""
    branch = ""
    commit = ""
    tag = ""
    any1 = ""
End Sub

'初期化
Public Sub Init(ByVal row As Integer)
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("params")
    
    Dim C0 As String: C0 = Define.TARGET_PARAM_CLM_0
    Dim C1 As String: C1 = Define.TARGET_PARAM_CLM_1
    Dim C2 As String: C2 = Define.TARGET_PARAM_CLM_2
    Dim C3 As String: C3 = Define.TARGET_PARAM_CLM_3
    Dim C4 As String: C4 = Define.TARGET_PARAM_CLM_4
    Dim C5 As String: C5 = Define.TARGET_PARAM_CLM_5
    Dim C6 As String: C6 = Define.TARGET_PARAM_CLM_6
    Dim C6a As String: C6a = Define.TARGET_PARAM_CLM_6a
    Dim C7 As String: C7 = Define.TARGET_PARAM_CLM_7
    Dim C8 As String: C8 = Define.TARGET_PARAM_CLM_8
    Dim C9 As String: C9 = Define.TARGET_PARAM_CLM_9
    Dim C10 As String: C10 = Define.TARGET_PARAM_CLM_10
    
    is_enable_s = sheet.Range(C0 & row).value
    vbprj_file_path = sheet.Range(C1 & row).value
    lot = sheet.Range(C2 & row).value
    phase = sheet.Range(C3 & row).value
    pid = sheet.Range(C4 & row).value
    ticket_num = sheet.Range(C5 & row).value
    step = sheet.Range(C6 & row).value
    dest_base_dir = sheet.Range(C6a & row).value
    branch = sheet.Range(C7 & row).value
    commit = sheet.Range(C8 & row).value
    tag = sheet.Range(C9 & row).value
    any1 = sheet.Range(C10 & row).value
End Sub

'検証
Public Sub Validate(ByVal row As Integer)
    '(0)
    If is_enable_s = "Enable" Then
        is_enable = True
    Else
        is_enable = False
    End If

    If is_enable = False Then
        Exit Sub
    End If

    '(1)
    If proc_type = PROCESS_TYPE.PROC_001 Or _
       proc_type = PROCESS_TYPE.PROC_002 Or _
       proc_type = PROCESS_TYPE.PROC_003 Then
        If vbprj_file_path = "" Then
            Err.Raise 53, , HEADER & "VBプロジェクトファイルが指定されていません(row=" & row & ")"
        End If
        
        If proc_type = PROCESS_TYPE.PROC_001 Or _
           proc_type = PROCESS_TYPE.PROC_002 Then
            If Common.IsExistsFile(vbprj_file_path) = False Then
                Err.Raise 53, , HEADER & "VBプロジェクトファイルが見つかりません(vbprj_file_path=" & vbprj_file_path & ", row=" & row & ")"
            End If
        End If
    
    End If
    
    If proc_type = PROCESS_TYPE.PROC_002 Or _
       proc_type = PROCESS_TYPE.PROC_003 Or _
       proc_type = PROCESS_TYPE.PROC_004 Or _
       proc_type = PROCESS_TYPE.PROC_005 Or _
       proc_type = PROCESS_TYPE.PROC_006 Or _
       proc_type = PROCESS_TYPE.DELETE_BRANCH Or _
       proc_type = PROCESS_TYPE.DELETE_TAG Then
       '(2)
        If lot = "" Then
            Err.Raise 53, , HEADER & "Lotが指定されていません(row=" & row & ")"
        End If
        
        '(3)
        If phase = "" Then
            Err.Raise 53, , HEADER & "Phaseが指定されていません(row=" & row & ")"
        End If
    End If
    
    If proc_type = PROCESS_TYPE.PROC_002 Or _
       proc_type = PROCESS_TYPE.PROC_003 Or _
       proc_type = PROCESS_TYPE.PROC_004 Or _
       proc_type = PROCESS_TYPE.PROC_005 Or _
       proc_type = PROCESS_TYPE.PROC_006 Or _
       proc_type = PROCESS_TYPE.DELETE_BRANCH Or _
       proc_type = PROCESS_TYPE.DELETE_TAG Then
        '(4)
        If pid = "" Then
            Err.Raise 53, , HEADER & "PIDが指定されていません(row=" & row & ")"
        End If
    End If
    
    If proc_type = PROCESS_TYPE.PROC_002 Or _
       proc_type = PROCESS_TYPE.PROC_003 Or _
       proc_type = PROCESS_TYPE.PROC_004 Or _
       proc_type = PROCESS_TYPE.PROC_005 Or _
       proc_type = PROCESS_TYPE.DELETE_BRANCH Then
        '(5)
        If ticket_num = "" Then
            Err.Raise 53, , HEADER & "ticket_numが指定されていません(row=" & row & ")"
        End If
    End If
    
    If proc_type = PROCESS_TYPE.PROC_002 Or _
       proc_type = PROCESS_TYPE.PROC_003 Or _
       proc_type = PROCESS_TYPE.PROC_004 Or _
       proc_type = PROCESS_TYPE.PROC_005 Or _
       proc_type = PROCESS_TYPE.PROC_006 Or _
       proc_type = PROCESS_TYPE.DELETE_TAG Then
        '(6)
        If step = "" Then
            Err.Raise 53, , HEADER & "stepが指定されていません(row=" & row & ")"
        End If
    End If
    
    If proc_type = PROCESS_TYPE.PROC_002 Then
        '(6a)
        If dest_base_dir = "" Then
            Err.Raise 53, , HEADER & "dest_base_dirが指定されていません(row=" & row & ")"
        End If
    End If
    
    '(7)
    If branch = "" Then
        Err.Raise 53, , HEADER & "branchが指定されていません(row=" & row & ")"
    End If
    
    '(8)
    If commit = "" Then
        Err.Raise 53, , HEADER & "commitが指定されていません(row=" & row & ")"
    End If

    '(9)
    If tag = "" Then
        Err.Raise 53, , HEADER & "tagが指定されていません(row=" & row & ")"
    End If
    
    If proc_type = PROCESS_TYPE.RENAME_TAG Then
        '(10)
        If any1 = "" Then
            Err.Raise 53, , HEADER & "Any1が指定されていません(row=" & row & ")"
        End If
    End If
End Sub

'=========================================
'Getter
'=========================================

Public Function GetProcessType() As PROCESS_TYPE
    GetProcessType = proc_type
End Function

Public Function IsEnable() As Boolean
    IsEnable = is_enable
End Function

Public Function GetVBPrjFilePath() As String
    GetVBPrjFilePath = vbprj_file_path
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

Public Function GetTicketNum() As String
    GetTicketNum = ticket_num
End Function

Public Function GetStep() As String
    GetStep = step
End Function

Public Function GetDestBaseDir() As String
    GetDestBaseDir = dest_base_dir
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

Public Function GetAny1() As String
    GetAny1 = any1
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = HEADER & _
          "is_enable=(" & is_enable_s & "), " & _
          "vbprj_file_path=(" & vbprj_file_path & "), " & _
          "lot=(" & lot & "), " & _
          "phase=(" & phase & "), " & _
          "pid=(" & pid & "), " & _
          "ticket_num=(" & ticket_num & "), " & _
          "step=(" & step & "), " & _
          "dest_base_dir=(" & dest_base_dir & "), " & _
          "branch=(" & branch & "), " & _
          "commit=(" & commit & "), " & _
          "tag=(" & tag & "), " & _
          "any1=(" & any1 & ")"

    GetAllValue = ret
End Function

'---------------------------
' Setter
'---------------------------
Public Sub SetProcessType(ByVal type_ As PROCESS_TYPE)
    proc_type = type_
End Sub

