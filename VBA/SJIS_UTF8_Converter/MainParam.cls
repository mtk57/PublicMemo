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

'MUST
Private src_dir_path As String

'WANT
Private dst_dir_path As String
Private conv_type As String
Private extension As String
Private is_contain_subdir As Boolean
Private is_backup As Boolean

Private is_contain_subdir_s As String
Private is_backup_s As String


'�R���X�g���N�^
Private Sub Class_Initialize()
    'MUST
    src_dir_path = ""
    
    'WANT
    dst_dir_path = ""
    conv_type = ""
    extension = ""
    is_contain_subdir = False
    is_backup = False
End Sub

'������
Public Function Init() As String
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Const CLM = "D"
    Const i = 6
    
    'MUST
    src_dir_path = main_sheet.Range(CLM & i).value
    
    'WANT
    conv_type = main_sheet.Range(CLM & i + 1).value
    is_contain_subdir_s = main_sheet.Range(CLM & i + 2).value
    extension = main_sheet.Range(CLM & i + 3).value
    dst_dir_path = main_sheet.Range(CLM & i + 7).value
    is_backup_s = main_sheet.Range(CLM & i + 8).value
    
    Init = Validate()
End Function

'����
Private Function Validate() As String
    'MUST
    If src_dir_path = "" Then
        Validate = "SRC�t�H���_�p�X���w�肳��Ă��܂���"
        Exit Function
    ElseIf Common.IsExistsFolder(src_dir_path) = False Then
        Validate = "�w�肳�ꂽSRC�t�H���_�����݂��܂��� (" & src_dir_path & ")"
        Exit Function
    End If

    'WANT
    If conv_type = "" Then
        conv_type = "SJIS��UTF8"
    End If
    
    If is_contain_subdir_s = "" Or _
       is_contain_subdir_s = "NO" Then
       is_contain_subdir = False
    Else
        is_contain_subdir = True
    End If
    
    If extension = "" Then
        conv_type = "*.*"
    End If
    
    If is_backup_s = "" Or _
       is_backup_s = "NO" Then
       is_backup = False
    Else
        is_backup = True
    End If

    Validate = ""
End Function

'=========================================
'Getter
'=========================================

'MUST

'-----------------------------------------
'SRC�t�H���_�p�X
'-----------------------------------------
Public Function GetSrcDirPath() As String
    GetSrcDirPath = src_dir_path
End Function



'WANT


'-----------------------------------------
'DST�t�H���_�p�X
'-----------------------------------------
Public Function GetDestDirPath() As String
    GetDestDirPath = dst_dir_path
End Function

'-----------------------------------------
'�ϊ����
'-----------------------------------------
Public Function GetConvertType() As String
    GetConvertType = conv_type
End Function

'-----------------------------------------
'�T�u�t�H���_�܂ނ�?
'-----------------------------------------
Public Function IsContainSubDir() As Boolean
    IsContainSubDir = is_contain_subdir
End Function

'-----------------------------------------
'�g���q
'-----------------------------------------
Public Function GetExtension() As String
    GetExtension = extension
End Function

'-----------------------------------------
'�o�b�N�A�b�v�t�@�C�����쐬���邩?
'-----------------------------------------
Public Function IsBackup() As Boolean
    IsBackup = is_backup
End Function

'-----------------------------------------
Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "src_dir_path=(" & src_dir_path & "), " & _
          "dst_dir_path=(" & dst_dir_path & "), " & _
          "conv_type=(" & conv_type & "), " & _
          "is_contain_subdir=(" & is_contain_subdir_s & "), " & _
          "extension=(" & extension & "), " & _
          "is_backup=(" & is_backup_s & ")"
    GetAllValue = ret
End Function


'=========================================
'Setter
'=========================================
'�Ȃ�
