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

Private dict_ As Dict


'�R���X�g���N�^
Private Sub Class_Initialize()
    Set dict_ = New Dict
End Sub

'������
Public Sub Init()
End Sub

'=========================================
'Setter
'=========================================

Public Sub SetRowData(ByVal prj_path As String, ByVal ref_path As String)
    dict_.AppendStringArray prj_path, ref_path
End Sub

'������
Public Sub SetDict(ByRef arg_dict As Variant)
    Set dict_ = arg_dict
End Sub

'=========================================
'Getter
'=========================================

Public Function GetRowCount() As Long
    GetRowCount = dict_.GetKeyCount()
End Function

Public Function GetPrjPathList() As Variant()
    GetPrjPathList = dict_.GetKeys()
End Function

Public Function GetRefPath(ByVal prj_path As String) As Variant
    GetRefPath = dict_.GetValue(prj_path)
End Function

Public Function GetAllRefCount() As Long
    GetAllRefCount = dict_.GetAllValueCount()
End Function

'=========================================
' Public I/F
'=========================================
Public Function IsExistKey(ByVal key As String) As Boolean
    IsExistKey = dict_.IsExists(key)
End Function

Public Function IsExistValue( _
    ByVal key As String, _
    ByVal val As String, _
    Optional ByVal is_match As Boolean = False _
) As Boolean

    IsExistValue = False

    Dim i As Long
    Dim values() As String
    
    If IsExistKey(key) = False Then
        IsExistValue = False
        Exit Function
    End If
    
    values = dict_.GetValue(key)

    For i = 0 To UBound(values)
        If is_match = False Then
            '���̂܂ܔ�r
            If values(i) = val Then
                IsExistValue = True
                Exit Function
            End If
        Else
            '�啶���������ɂ���ʂ����ɔ�r
            If UCase(values(i)) = UCase(val) Then
                IsExistValue = True
                Exit Function
            End If
        End If
    Next i
End Function

'Public Function DeepCopy() As RefFiles
'
'    Dim ret As RefFiles
'    Set ret = New RefFiles
'
'    Dim i As Long
'    Dim j As Long
'
'    For i = 0 To GetRowCount()
'        Dim key As String: key = GetPrjPathList()(i)
'
'        Dim values() As String
'        values = GetRefPath(key)
'
'        For j = 0 To UBound(values)
'            Dim value As String: value = values(j)
'            ret.SetRowData key, value
'        Next j
'    Next i
'
'    DeepCopy = ret
'
'End Function
