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
Private ignore_keys As String
Private ignore_key_list() As String

Private Const CLM_ = "O"
Private Const ROW_ = 6

'�R���X�g���N�^
Private Sub Class_Initialize()
    src_dir_path = ""
    ignore_keys = ""
    Erase ignore_key_list
End Sub

'������
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    src_dir_path = main_sheet.Range(CLM_ & ROW_).value
    ignore_keys = main_sheet.Range(CLM_ & ROW_ + 2).value
End Sub

'����
Public Sub Validate()
    If src_dir_path = "" Then
        Err.Raise 53, , "VBP�t�@�C������������t�H���_�p�X���w�肳��Ă��܂���"
    ElseIf Common.IsExistsFolder(src_dir_path) = False Then
        Err.Raise 53, , "VBP�t�@�C������������t�H���_�����݂��܂��� (" & src_dir_path & ")"
    End If

    '���O�L�[���X�g���쐬
    ignore_key_list = Split(ignore_keys, ",")
    
End Sub

'=========================================
'Getter
'=========================================

Public Function GetSrcDirPath() As String
    GetSrcDirPath = src_dir_path
End Function

Public Function GetIgnoreKeys() As String
    GetIgnoreKeys = ignore_keys
End Function

Public Function IsExistIgnoreKey(ByVal key As String) As Boolean
    IsExistIgnoreKey = False

    If Common.IsEmptyArray(ignore_key_list) = True Then
        Exit Function
    End If
        
    Dim i As Integer
    For i = LBound(ignore_key_list) To UBound(ignore_key_list)
        If InStr(LCase(key), LCase(ignore_key_list(i))) > 0 Then
            '���O�L�[���܂ނ̂Ŗ���
            Exit Function
        End If
    Next i
    
    IsExistIgnoreKey = True
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "src_dir_path=(" & src_dir_path & "), " & _
          "ignore_keys=(" & ignore_keys & ")"

    GetAllValue = ret
End Function