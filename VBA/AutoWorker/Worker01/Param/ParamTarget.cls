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

'�����o�[�ϐ�
Private vbprj_file_path As String
Private pid As String
Private ticket_num As String
Private branch As String
Private commit As String
Private tag As String

'�R���X�g���N�^
Private Sub Class_Initialize()
    HEADER = "[Target Param] "
    
    vbprj_file_path = ""
    pid = ""
    ticket_num = ""
    branch = ""
    commit = ""
    tag = ""
End Sub

'������
Public Sub Init(ByVal row As Integer)
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(Define.SHEET_01)
    
    Dim C1 As String: C1 = Define.TARGET_PARAM_CLM_1
    Dim C2 As String: C2 = Define.TARGET_PARAM_CLM_2
    Dim C3 As String: C3 = Define.TARGET_PARAM_CLM_3
    Dim C4 As String: C4 = Define.TARGET_PARAM_CLM_4
    Dim C5 As String: C5 = Define.TARGET_PARAM_CLM_5
    Dim C6 As String: C6 = Define.TARGET_PARAM_CLM_6
    
    vbprj_file_path = main_sheet.Range(C1 & row).value
    pid = main_sheet.Range(C2 & row).value
    ticket_num = main_sheet.Range(C3 & row).value
    branch = main_sheet.Range(C4 & row).value
    commit = main_sheet.Range(C5 & row).value
    tag = main_sheet.Range(C6 & row).value
End Sub

'����
Public Sub Validate(ByVal row As Integer)
    If vbprj_file_path = "" Then
        Err.Raise 53, , HEADER & "VB�v���W�F�N�g�t�@�C�����w�肳��Ă��܂���(row=" & row & ")"
    End If
    
    If Common.IsExistsFile(vbprj_file_path) = False Then
        Err.Raise 53, , HEADER & "VB�v���W�F�N�g�t�@�C����������܂���(vbprj_file_path=" & vbprj_file_path & ", row=" & row & ")"
    End If
    
    If pid = "" Then
        Err.Raise 53, , HEADER & "pid���w�肳��Ă��܂���(row=" & row & ")"
    End If
    
    If ticket_num = "" Then
        Err.Raise 53, , HEADER & "ticket_num���w�肳��Ă��܂���(row=" & row & ")"
    End If
    
    If branch = "" Then
        Err.Raise 53, , HEADER & "branch���w�肳��Ă��܂���(row=" & row & ")"
    End If
    
    If commit = "" Then
        Err.Raise 53, , HEADER & "commit���w�肳��Ă��܂���(row=" & row & ")"
    End If

    If tag = "" Then
        Err.Raise 53, , HEADER & "tag���w�肳��Ă��܂���(row=" & row & ")"
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


