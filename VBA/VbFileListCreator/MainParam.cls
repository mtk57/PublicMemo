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

Private src_dir_path As String

Private Const CLM_ = "O"
Private Const ROW_ = 6

'�R���X�g���N�^
Private Sub Class_Initialize()
    src_dir_path = ""
End Sub

'������
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    src_dir_path = main_sheet.Range(CLM_ & ROW_).value
End Sub

'����
Public Sub Validate()
    If src_dir_path = "" Then
        Err.Raise 53, , "��������t�H���_�p�X���w�肳��Ă��܂���"
    ElseIf Common.IsExistsFolder(src_dir_path) = False Then
        Err.Raise 53, , "��������t�H���_�����݂��܂��� (" & src_dir_path & ")"
    End If
End Sub

'=========================================
'Getter
'=========================================

'-----------------------------------------
'��������t�H���_�p�X
'-----------------------------------------
Public Function GetSrcDirPath() As String
    GetSrcDirPath = src_dir_path
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "src_dir_path=(" & src_dir_path & ")"

    GetAllValue = ret
End Function