VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RefFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�����o�[�ϐ�

Private vbp_file_path As String
Private ref_file_path() As String

'�R���X�g���N�^
Private Sub Class_Initialize()
    vbp_file_path = ""
    Erase ref_file_path
End Sub

'������
Public Sub Init()

End Sub

'=========================================
'Setter
'=========================================

'-----------------------------------------
'VBP�t�@�C���p�X
'-----------------------------------------
Public Sub SetSrcDirPath(ByVal path As String)
    vbp_file_path = path
End Sub

'-----------------------------------------
'VBP���Q�Ƃ��Ă���t�@�C���p�X
'-----------------------------------------
Public Sub AppendRefFilePath(ByVal path As String)
    Common.AppendArray ref_file_path, path
End Sub

'=========================================
'Getter
'=========================================

'-----------------------------------------
'VBP�t�@�C���p�X
'-----------------------------------------
Public Function GetSrcDirPath() As String
    GetSrcDirPath = vbp_file_path
End Function

'-----------------------------------------
'VBP���Q�Ƃ��Ă���t�@�C���p�X���X�g�̗v�f��
'-----------------------------------------
Public Function GetAppendRefFileCount() As Long
    GetAppendRefFileCount = UBound(ref_file_path)
End Function

'-----------------------------------------
'VBP���Q�Ƃ��Ă���t�@�C���p�X���X�g
'-----------------------------------------
Public Function GetRefFile(ByVal index As Long) As String
    GetRefFile = ref_file_path(index)
End Function

'-----------------------------------------
'VBP���Q�Ƃ��Ă���t�@�C���p�X���X�g
'-----------------------------------------
Public Function GetRefFiles() As String()
    GetRefFiles = ref_file_path
End Function
