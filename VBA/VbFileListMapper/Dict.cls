VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const VERSION = "1.0.2"

Private dict_ As Object

'�R���X�g���N�^
Private Sub Class_Initialize()
  
  Set dict_ = CreateObject("Scripting.Dictionary")
  
End Sub

'�L�[�̑��ݗL����Ԃ�
Public Function IsExists(ByVal key As Variant) As Boolean
    IsExists = dict_.Exists(key)
End Function

'�L�[�ɕR�Â��l��ǉ�����
Public Sub AppendValue(ByVal key As Variant, ByVal val As Variant)
    If IsNull(key) Then
        Err.Raise 53, , "�L�[����ł�"
    End If
    
    If dict_.Exists(key) = False Then
        dict_.Add key, val
    Else
        dict_(key) = val
    End If
End Sub

'�L�[�ɕR�Â�String�z��ɒl��ǉ�����
Public Sub AppendStringArray(ByVal key As Variant, ByVal val As String)
    If IsNull(key) Then
        Err.Raise 53, , "�L�[����ł�"
    End If

    Dim i As Long
    Dim wk_ary() As String
    
    If dict_.Exists(key) = False Then
        ReDim Preserve wk_ary(0)
        wk_ary(0) = val
    
        dict_.Add key, wk_ary
    Else
        Dim cnt As Long
        cnt = UBound(dict_(key)) + 1
        
        wk_ary = dict_(key)
        
        ReDim Preserve wk_ary(cnt)
        
        wk_ary(cnt) = val
        
        dict_(key) = wk_ary
    End If
End Sub

'�L�[�ɕR�Â��l��Ԃ�
Public Function GetValue(ByVal key As Variant) As Variant
    If IsNull(key) Then
        Err.Raise 53, , "�L�[����ł�"
    End If
    
    GetValue = dict_(key)
End Function

'�L�[����Ԃ�
Public Function GetKeyCount() As Long
    GetKeyCount = UBound(dict_.Keys)
End Function

'�L�[�̈ꗗ��Ԃ�
Public Function GetKeys() As Variant()
    GetKeys = dict_.Keys
End Function

'�l�̑�����Ԃ�
Public Function GetAllValueCount() As Long
    Dim ret As Long
    
    Dim i As Long
    Dim key As String
    Dim values() As String
    
    For i = 0 To UBound(dict_.Keys)
        key = GetKeys()(i)
        values = dict_(key)
        
        ret = ret + UBound(values) + 1
    Next i

    GetAllValueCount = ret - 1
End Function

'String�z��̃R�s�[��Ԃ�(�{�c)
'Public Function CopyStringArray(ByRef dst_dict As Dict) As Boolean
'    Set dst_dict = New Dict
'
'    Dim i As Long
'    Dim j As Long
'
'    For i = 0 To GetKeyCount()
'        Dim key As String: key = GetKeys()(i)
'
'        Dim values() As String
'        values = GetValue(key)
'
'        For j = 0 To UBound(values)
'            Dim value As String: value = values(j)
'            dst_dict.AppendStringArray key, value
'        Next j
'    Next i
'    CopyStringArray = True
'End Function

