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

Private src_path As String
Private dst_path As String

Private Const CLM_ = "L"
Private Const ROW_ = 7

'�R���X�g���N�^
Private Sub Class_Initialize()
    src_path = ""
    dst_path = ""
End Sub

'������
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    src_path = main_sheet.Range(CLM_ & ROW_).value
    dst_path = main_sheet.Range(CLM_ & ROW_ + 1).value
End Sub

'����
Public Sub Validate()
    If src_path = "" Then
        Err.Raise 53, , "�t�@�C���p�X���w�肳��Ă��܂���"
    ElseIf Common.IsExistsFile(src_path) = False Then
        Err.Raise 53, , "�t�@�C�������݂��܂��� (" & src_path & ")"
    End If
End Sub

'=========================================
'Getter
'=========================================

Public Function GetSrcPath() As String
    GetSrcPath = src_path
End Function

Public Function GetDstPath() As String
    GetDstPath = dst_path
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "src_path=(" & src_path & ", )" & _
          "dst_path=(" & dst_path & ")"

    GetAllValue = ret
End Function