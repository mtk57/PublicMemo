VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParseResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�����o�[�ϐ�
Private target As String
Private format_type As String
Private file_path As String
Private extension As String
Private row_num As Long
Private clm_num As Long
Private encode As String
Private raw_line As String
Private begin_codes() As String
Private begin_class As String
Private begin_members() As String

'�R���X�g���N�^
Private Sub Class_Initialize()
    target = ""
    format_type = ""
    file_path = ""
    extension = ""
    row_num = 0
    clm_num = 0
    encode = ""
    raw_line = ""
    begin_class = ""
    Erase begin_codes
    Erase begin_members
End Sub

'������
Public Sub Init(ByVal target_ As String, ByVal format_type_ As String)
    target = target_
    format_type = format_type_
    
    Parse
End Sub

Private Sub Parse()
    If format_type = "sakura" Then
        ParseForSakura
    End If
End Sub

'target�ɑz�肷�镶����͈ȉ��Ƃ���
'"C:\_git\PublicMemo\VBA\BeginParser\testdata\src\begin_sample.bas(2,4)  [UTF-8]:     Begin MyObject"
'
'�܂�A�ȉ��̍\���ł��邱�ƁB
'�t�@�C���̐�΃p�X        :"C:\_git\PublicMemo\VBA\BeginParser\testdata\src\begin_sample.bas"
'(�s�ԍ�, ��ԍ�)          :"(2,4)"
'���p�X�y�[�X*2            :"  "
'�G���R�[�h(��:[UTF-8])    :"[UTF-8]"
'���p�R����                :":"
'GREP����                  :"     Begin MyObject""
Private Sub ParseForSakura()
    Dim wk() As String
    Dim path_wk As String
    Dim pos_wk As String
    Dim pos_list() As String
    Dim i As Long
    
    '�������񔼊pSP�ŕ���
    wk = Split(target, " ")
    wk = Common.DeleteEmptyArray(wk)

    '�t�@�C���p�X���擾
    path_wk = wk(0)
    file_path = Mid(path_wk, 1, InStrRev(path_wk, "(") - 1)
    
    '�g���q���擾
    extension = Common.GetFileExtension(file_path)
    
    '��������"("����")"�܂ł��擾
    pos_wk = Mid(path_wk, InStrRev(path_wk, "("), Len(path_wk))
    
    '�s�E����擾
    pos_list = Common.DeleteEmptyArray(Common.GetMatchByRegExp(pos_wk, "\d*", True))
    If UBound(pos_list) < 1 Then
        Err.Raise 53, , "GREP���ʂ̍s��t�H�[�}�b�g���s���ł� (target=" & target & ")"
    End If
    row_num = pos_list(0)
    clm_num = pos_list(1)
    
    '�G���R�[�h���擾
    For i = 0 To UBound(wk)
        If Common.IsMatchByRegExp(wk(i), "\[.*\]", True) = True Then
            encode = Replace(Replace(wk(i), "[", ""), "]:", "")
            Exit For
        End If
    Next i
    
    If encode = "" Then
        Err.Raise 53, , "GREP���ʂ���G���R�[�h�����񂪌�����܂��� (target=" & target & ")"
    ElseIf encode <> "SJIS" And encode <> "UTF-8" Then
        Err.Raise 53, , "���T�|�[�g�̃G���R�[�h�ł� (target=" & target & ")"
    End If
    
    'GREP���ʂ��擾
    raw_line = Mid(target, InStr(target, "]:") + 2, Len(target))
    
End Sub

'--------------------
'Getter
'--------------------

Public Function GetTarget() As String
    GetTarget = target
End Function

Public Function GetFormatType() As String
    GetFormatType = format_type
End Function

Public Function GetFilePath() As String
    GetFilePath = file_path
End Function

Public Function GetExtension() As String
    GetExtension = extension
End Function

Public Function GetRowNum() As Long
    GetRowNum = row_num
End Function

Public Function GetClmNum() As Long
    GetClmNum = clm_num
End Function

Public Function GetEncode() As String
    GetEncode = encode
End Function

Public Function GetRawLine() As String
    GetRawLine = raw_line
End Function

Public Function GetBeginCodes() As String()
    GetBeginCodes = begin_codes
End Function

Public Function GetBeginCodesCount() As Long
    If Common.IsEmptyArray(begin_codes) = True Then
        GetBeginCodesCount = 0
        Exit Function
    End If
    
    Dim cnt As Long: cnt = UBound(begin_codes)
    cnt = cnt + 1
    GetBeginCodesCount = cnt
End Function

Public Function GetBeginClass() As String
    If begin_class <> "" Then
        GetBeginClass = begin_class
        Exit Function
    End If
    
    If Common.IsEmptyArray(begin_codes) = True Then
        Err.Raise 53, , "[GetBeginClass] begin_codes����ł�"
    End If
    
    Dim wk As String: wk = begin_codes(0)
    If Left(wk, 6) <> "Begin " Then
        Err.Raise 53, , "[GetBeginClass] Begin��������܂���"
    End If
    
    begin_class = Trim(Replace(wk, "Begin ", ""))

    GetBeginClass = begin_class
End Function

Public Function GetBeginMembers() As String()
    GetBeginMembers = begin_members
End Function

Public Function GetBeginMembersCount() As Long
    If Common.IsEmptyArray(begin_members) = True Then
        GetBeginMembersCount = 0
        Exit Function
    End If
    
    Dim cnt As Long: cnt = UBound(begin_members)
    cnt = cnt + 1
    GetBeginMembersCount = cnt
End Function

'--------------------
'Setter
'--------------------
Public Sub SetBeginCodes(ByRef codes() As String)
    begin_codes = codes
End Sub

Public Sub SetBeginClass(ByVal class_name As String)
    begin_class = class_name
End Sub

Public Sub SetBeginMembers(ByRef members() As String)
    begin_members = members
End Sub
