Attribute VB_Name = "Common"
Option Explicit

Private Const VERSION = "1.5.7"

Public Type MethodInfoStruct
    Raw As String
    Name As String
    ret As String
    params() As String
End Type

Public Type GrepResultInfoStruct
    ResultRaw As String
    FilePath As String
    Ext As String
    row As Long
    Clm As Long
    Contents As String
    IsError As Boolean
    ErrorInfo As String
    MethodInfo As MethodInfoStruct
End Type

Public Enum GrepAppEnum
    sakura
    hidemaru
End Enum

Public Enum LangEnum
    VB6
    VBNET
    CSharp
End Enum

Private Declare PtrSafe Function GetPrivateProfileString Lib _
    "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
) As Long

Private Declare PtrSafe Function WritePrivateProfileString Lib _
    "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String _
) As Long

Private Declare PtrSafe Sub GetLocalTime Lib _
    "kernel32" ( _
    lpSystemTime As SYSTEMTIME _
)

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'���O�t�@�C���ԍ�
Private logfile_num As Integer
Private is_log_opened As Boolean

Private Const GIT_BASH = "C:\Program Files\Git\usr\bin\bash.exe"

Public Function IsVariantNull(v As Variant) As Boolean
    On Error Resume Next
    IsVariantNull = IsEmpty(v)
    On Error GoTo 0
End Function

'-------------------------------------------------------------
' �z����N�C�b�N�\�[�g�ŏ����\�[�g����
' argAry : I/O : �z��
' lngMin : I : �\�[�g�͈�(�ŏ�)  ���C��
' lngMax : I : �\�[�g�͈�(�ő�)  ���C��
' keyPos : I : �\�[�g�L�[�C���f�N�X  ���C��
' BASE : https://excel-ubara.com/excelvba5/EXCELVBA229.html
'-------------------------------------------------------------
Public Sub QuickSortArray( _
    ByRef argAry() As Variant, _
    Optional ByVal lngMin As Long = -1, _
    Optional ByVal lngMax As Long = -1, _
    Optional ByVal keyPos As Long = 0 _
)
    If IsEmptyArray(argAry) = True Then
        Exit Sub
    End If
    
    If lngMin = -1 Then
        lngMin = LBound(argAry)
    End If
    If lngMax = -1 Then
        lngMax = UBound(argAry)
    End If

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim vBase As Variant
    Dim vSwap As Variant
    
    vBase = argAry(Int((lngMin + lngMax) / 2), keyPos)
    i = lngMin
    j = lngMax
    
    Do
        Do While argAry(i, keyPos) < vBase
            i = i + 1
        Loop
        
        Do While argAry(j, keyPos) > vBase
            j = j - 1
        Loop
        
        If i >= j Then Exit Do
        
        For k = LBound(argAry, 2) To UBound(argAry, 2)
            vSwap = argAry(i, k)
            argAry(i, k) = argAry(j, k)
            argAry(j, k) = vSwap
        Next
        i = i + 1
        j = j - 1
    Loop
    
    If (lngMin < i - 1) Then
        Call QuickSortArray(argAry, lngMin, i - 1, keyPos)
    End If
    
    If (lngMax > j + 1) Then
        Call QuickSortArray(argAry, j + 1, lngMax, keyPos)
    End If
    
End Sub

'-------------------------------------------------------------
' �z��̎w��ʒu�Ɏw��T�C�Y�̋�s��}������
'  ��:["A", "B", "C"]�Ƃ����z��̏ꍇ�A�ȉ��̂悤�ɂȂ�
'    pos=2, size=2:["A", "B", "", "", "C"]
'    pos=0, size=1:["", "A", "B", "C"]
' ary : I : String�z��
' pos : I : �}���ʒu(0=�擪)
' size : I : �}���s��
' RET : �}�����String�z��
'-------------------------------------------------------------
Public Function InsertEmptyRows(ByRef ary() As String, ByVal pos As Long, ByVal size As Long) As String()
    '�����`�F�b�N
    If Common.IsEmptyArray(ary) = True Then
        InsertEmptyRows = ary
        Exit Function
    End If
    
    If UBound(ary) < pos Then
        Err.Raise 53, , "[InsertEmptyRows] pos���z��T�C�Y�𒴂��Ă��܂�! (pos=" & pos & ")"
    End If
    
    If pos < 0 Then
        pos = 0
    End If
    
    If size < 1 Then
        InsertEmptyRows = ary
        Exit Function
    End If

    Dim i As Long
    Dim ret() As String
    ReDim ret(LBound(ary) To UBound(ary) + size)
    
    For i = LBound(ret) To UBound(ret)
        If i < pos Then
            ret(i) = ary(i)
        ElseIf i >= pos And i < pos + size Then
            ret(i) = ""
        Else
            ret(i) = ary(i - size)
        End If
    Next i
    
    InsertEmptyRows = ret
End Function

'-------------------------------------------------------------
' �w�肵���s�̎��̍s����w�肵���s���������s��ǉ�����
' ws : I : ���[�N�V�[�g
' startRowNum : I : �w��s (1�`)
' insertRowCnt : I : �ǉ��s�� (1�`)
'-------------------------------------------------------------
Public Sub InsertRows(ByRef ws As Worksheet, ByVal startRowNum As Long, ByVal insertRowCnt As Long)
    '�����`�F�b�N
    If startRowNum < 1 Then
        startRowNum = 1
    End If
    
    If insertRowCnt < 1 Then
        Exit Sub
    End If
    
    ws.Rows(startRowNum + 1 & ":" & startRowNum + insertRowCnt).Insert shift:=xlDown
End Sub

'-------------------------------------------------------------
' Grep���ʂ���͂��ă��\�b�h����Ԃ�
' grepResults : I :Grep���ʂ̔z��
'                   <sakura�̏ꍇ>
'                   �t�@�C���t���p�X + (�s�ԍ�, ��ԍ�) + " " + [�G���R�[�h] + ":" + ���e
' grepApp : I : Grep�����A�v�� (�����_�ł�sakura�̂݃T�|�[�g)  ���C��
' lang : I : Grep�������� (�����_�ł�VB6�̂݃T�|�[�g)  ���C��
' Ret : ���\�b�h��� (Grep���ʂ̃R�[�h�������Ă��郁�\�b�h)
'-------------------------------------------------------------
Public Function GetMethodInfoFromGrepResult( _
    ByRef grepResults() As String, _
    Optional ByVal grepApp As GrepAppEnum = GrepAppEnum.sakura, _
    Optional ByVal lang As LangEnum = LangEnum.VB6 _
) As GrepResultInfoStruct()

    Dim ret() As GrepResultInfoStruct

    '�����`�F�b�N
    If Common.IsEmptyArray(grepResults) = True Then
         GetMethodInfoFromGrepResult = ret
         Exit Function
    End If
    
    If grepApp <> GrepAppEnum.sakura Then
        GetMethodInfoFromGrepResult = ret
        Exit Function
    End If

    If lang <> LangEnum.VB6 Then
        GetMethodInfoFromGrepResult = ret
        Exit Function
    End If
    

    ReDim Preserve ret(UBound(grepResults))

    Dim i As Long
    Dim raw_contents() As String
    Dim before_path As String: before_path = ""
    
    For i = 0 To UBound(grepResults)
        Dim cur_info As GrepResultInfoStruct
        
        'Grep�����擾����
        cur_info = GetGrepInfo(grepResults(i), grepApp, lang)
        
        '�R�����g�̏ꍇ�͖�������
        If IsCommentCode(cur_info.Contents, cur_info.Ext) Then
            GoTo CONTINUE_A
        End If
        
        If before_path = cur_info.FilePath Then
            '�O��Ɠ����t�@�C���Ȃ̂Ń�������̓��e���g���܂킷
        Else
            '�O��ƈقȂ�t�@�C���ɐ؂�ւ�����̂Ńt�@�C���̓��e��ǂݍ���
            If IsExistsFile(cur_info.FilePath) = False Then
                '�t�@�C�������݂��Ȃ��̂Ŗ�������
                cur_info.IsError = True
                cur_info.ErrorInfo = "��File not exist.(" & cur_info.FilePath & ")"
                GoTo CONTINUE_A
            End If
            
            raw_contents = GetContents(cur_info.FilePath)
            before_path = cur_info.FilePath
        End If
        

        '�s�ԍ��̈ʒu�����Ɍ������ă��\�b�h��T��
        cur_info.MethodInfo = FindMethodByGrepResultInfo(cur_info, raw_contents, grepApp, lang)
        
CONTINUE_A:

        ret(i) = cur_info
        
    Next i
    
    
    GetMethodInfoFromGrepResult = ret

End Function

'-------------------------------------------------------------
' Grep���ʂ���͂���GREP����Ԃ�
' grepResult : I :Grep����
'                   <sakura�̏ꍇ>
'                   �t�@�C���t���p�X + (�s�ԍ�, ��ԍ�) + " " + [�G���R�[�h] + ":" + ���e
' grepApp : I : Grep�����A�v�� (�����_�ł�sakura�̂݃T�|�[�g)  ���C��
' lang : I : Grep�������� (�����_�ł�VB6�̂݃T�|�[�g)  ���C��
' Ret : GREP���
'-------------------------------------------------------------
Public Function GetGrepInfo( _
    ByVal grepResult As String, _
    Optional ByVal grepApp As GrepAppEnum = GrepAppEnum.sakura, _
    Optional ByVal lang As LangEnum = LangEnum.VB6 _
) As GrepResultInfoStruct

    Dim ret As GrepResultInfoStruct

    '�����`�F�b�N
    If grepResult = "" Then
        GetGrepInfo = ret
        Exit Function
    End If
    
    If grepApp <> GrepAppEnum.sakura Then
        GetGrepInfo = ret
        Exit Function
    End If

    If lang <> LangEnum.VB6 Then
        GetGrepInfo = ret
        Exit Function
    End If

    Dim regex_result() As String
    regex_result = GetMatchByRegExp(grepResult, "^.*\(\d+,\d+\)", False)
    
    If IsEmptyArray(regex_result) Then
        GetGrepInfo = ret
        Exit Function
    End If
    
    ret.ResultRaw = grepResult
    ret.FilePath = ReplaceByRegExp(regex_result(0), "\(\d+,\d+\)", "", False)
    ret.Ext = GetFileExtension(ret.FilePath)
    Dim wk As String: wk = Trim(Replace(Replace(GetMatchByRegExp(regex_result(0), "\(\d+,\d+\)", False)(0), "(", ""), ")", ""))
    ret.row = CLng(Split(wk, ",")(0))
    ret.Clm = CLng(Split(wk, ",")(1))
    ret.Contents = ReplaceByRegExp(grepResult, "^.*\]:", "", False)
    
    GetGrepInfo = ret
End Function

'-------------------------------------------------------------
'�Ώۃt�@�C����ǂݍ���œ��e��z��ŕԂ�
' path : I : �Ώۃt�@�C���p�X
' Ret : �ǂݍ��񂾓��e
'-------------------------------------------------------------
Public Function GetContents(ByVal path As String) As String()
    Dim ret() As String
    
    If IsExistsFile(path) = False Then
        GetContents = ret
        Exit Function
    End If
    
    If IsUTF8(path) = False Then
        GetContents = Split(ReadTextFileBySJIS(path), vbCrLf)
    Else
        GetContents = Split(Replace(ReadTextFileByUTF8(path), vbCrLf, vbLf), vbLf)
    End If

End Function

'-------------------------------------------------------------
' GREP���ʏ�񂩂烁�\�b�h��T���ă��\�b�h���Ƃ��ĕԂ�
' grepResult : I : GREP���ʏ��
' contents : I : �T���Ώۃt�@�C���̓��e
' grepApp : I : Grep�����A�v�� (�����_�ł�sakura�̂݃T�|�[�g)  ���C��
' lang : I : Grep�������� (�����_�ł�VB6�̂݃T�|�[�g)  ���C��
' Ret : ���\�b�h���
'-------------------------------------------------------------
Public Function FindMethodByGrepResultInfo( _
    ByRef grepResult As GrepResultInfoStruct, _
    ByRef Contents() As String, _
    Optional ByVal grepApp As GrepAppEnum = GrepAppEnum.sakura, _
    Optional ByVal lang As LangEnum = LangEnum.VB6 _
) As MethodInfoStruct

    Dim ret As MethodInfoStruct
    
    '�����`�F�b�N
    If Common.IsEmptyArray(Contents) = True Then
         FindMethodByGrepResultInfo = ret
         Exit Function
    End If
    
    If grepApp <> GrepAppEnum.sakura Then
        FindMethodByGrepResultInfo = ret
        Exit Function
    End If

    If lang <> LangEnum.VB6 Then
        FindMethodByGrepResultInfo = ret
        Exit Function
    End If
    
    
    If lang = LangEnum.VB6 Or lang = VBNET Then
        Dim method_type As String
        Dim method_start_row As Long
    
        '���\�b�h(Function/Sub)�̃^�C�v�������� (�����s�͍l�����Ȃ�)
        method_type = FindMethodTypeForVB(Contents, grepResult.row)
        
        If method_type = "" Then
            '�����ł����B�B
            FindMethodByGrepResultInfo = ret
            Exit Function
        End If
        
        '���\�b�h(Function/Sub)�̊J�n�s�ԍ���������
        method_start_row = FindMethodStartRowForVB(Contents, grepResult.row, method_type)
        
        If method_start_row = -1 Then
            '�����ł����B�B
            FindMethodByGrepResultInfo = ret
            Exit Function
        End If
        
        '���\�b�h(Function/Sub)�̊J�n�s�ԍ����烁�\�b�h�����擾����
        ret = GetMethodInfoForVB(Contents, method_start_row, method_type)
        
    End If
    
    FindMethodByGrepResultInfo = ret

End Function

'-------------------------------------------------------------
' VB�̃��\�b�h(Function/Sub)�̎�ނ�������
' (�����s�͍l�����Ȃ�)
' contents : I : �T���Ώۃt�@�C���̓��e
' startRow : I : �J�n�s
' Ret : ����(Function/Sub), �����ł����B("")
'-------------------------------------------------------------
Public Function FindMethodTypeForVB( _
    ByRef Contents() As String, _
    ByVal startRow As Long _
) As String

    Dim ret As String: ret = ""

    '�����`�F�b�N
    If Common.IsEmptyArray(Contents) = True Then
         FindMethodTypeForVB = ret
         Exit Function
    End If

    If startRow < 0 Then
         FindMethodTypeForVB = ret
         Exit Function
    End If
    
    Dim i As Long
    
    For i = 0 To UBound(Contents)
        If (startRow + i) > UBound(Contents) Then
            Exit For
        End If
        
        Dim line As String: line = Contents(startRow + i)
        
        If IsCommentCode(line, "bas") Then
            '�R�����g�s�͖���
            GoTo CONTINUE
        End If
        
        If IsMatchByRegExp(line, "End\s+(Function|Sub)\b", True) = False Then
            '������Ȃ�
            GoTo CONTINUE
        End If
        
        '����
        ret = Trim(Replace(line, "End", ""))
        Exit For
        
CONTINUE:
        
    Next i
    
    FindMethodTypeForVB = ret

End Function

'-------------------------------------------------------------
' VB�̃��\�b�h(Function/Sub)�̊J�n�s�ԍ���������
' (�����A�߂�l�ȊO�̕����s�͍l�����Ȃ�)
' contents : I : �T���Ώۃt�@�C���̓��e
' startRow : I : �J�n�s
' methodType : I : ���(Function/Sub)
' Ret : ����(�J�n�s�ԍ�), �����ł����B(-1)
'-------------------------------------------------------------
Public Function FindMethodStartRowForVB( _
    ByRef Contents() As String, _
    ByVal startRow As Long, _
    ByVal methodType As String _
) As Long

    Dim ret As Long: ret = -1

    '�����`�F�b�N
    If Common.IsEmptyArray(Contents) = True Then
         FindMethodStartRowForVB = ret
         Exit Function
    End If

    If startRow < 0 Then
         FindMethodStartRowForVB = ret
         Exit Function
    End If
    
    Dim i As Long
    
    For i = 0 To UBound(Contents)
        If (startRow - i) < 0 Then
            Exit For
        End If
    
        Dim line As String: line = Contents(startRow - i)
        
        If IsCommentCode(line, "bas") Then
            '�R�����g�s�͖���
            GoTo CONTINUE
        End If
        
        If IsMatchByRegExp(line, "(Function|Sub)\s+[A-Za-z_][A-Za-z0-9_]*\(", True) = False Then
            '������Ȃ�
            GoTo CONTINUE
        End If
        
        '����
        ret = startRow - i
        Exit For
        
CONTINUE:
        
    Next i
    
    FindMethodStartRowForVB = ret

End Function

'-------------------------------------------------------------
' VB�̃��\�b�h(Function/Sub)����Ԃ�
' contents : I : �T���Ώۃt�@�C���̓��e
' startRow : I : ���\�b�h�J�n�s
' methodType : I : ���(Function/Sub)
' Ret : ���\�b�h���
'-------------------------------------------------------------
Public Function GetMethodInfoForVB( _
    ByRef Contents() As String, _
    ByVal startRow As Long, _
    ByVal methodType As String _
) As MethodInfoStruct

    Dim ret As MethodInfoStruct

    '�����`�F�b�N
    If Common.IsEmptyArray(Contents) = True Then
         GetMethodInfoForVB = ret
         Exit Function
    End If

    If startRow < 0 Then
         GetMethodInfoForVB = ret
         Exit Function
    End If
    
    If methodType <> "Function" And methodType <> "Sub" Then
         GetMethodInfoForVB = ret
         Exit Function
    End If
    
    Dim i As Long
    Dim merge_lines As String
    
    Dim start_clm As Long
    Dim end_clm As Long
    
    '�������񕡐��s��1�s�ɂ܂Ƃ߂�
    For i = 0 To UBound(Contents)
        If (startRow + i) > UBound(Contents) Then
            Exit For
        End If
    
        Dim line As String: line = Contents(startRow + i)
        
        If IsCommentCode(line, "bas") Then
            '�R�����g�s�͖���
            GoTo CONTINUE
        End If
        
        '�R�����g�폜
        line = Trim(RemoveRightComment(line, "bas"))
        
        '"_"�폜
        line = Trim(ReplaceByRegExp(line, "_$", "", False))
    
        '�}�[�W
        merge_lines = merge_lines + line
        
        
        '�����̏I���̊��ʂ�T��
        end_clm = FindMatchingBracketPositionForVB(merge_lines)
        
        If end_clm = 0 Then
            '������Ȃ�
            GoTo CONTINUE
        End If
        
        '��������
        
        If methodType = "Sub" Then
            Exit For
        End If
        
        'Function�̏ꍇ�͖߂�l�܂Ń}�[�W����Ă��邩�m�F����
        If IsMatchByRegExp(merge_lines, "Function\s+[A-Za-z_][A-Za-z0-9_]*\(.*\)(\s+As\s+[A-Za-z_][A-Za-z0-9_]*\(*\)*)*$", True) = True Then
            Exit For
        End If
        
CONTINUE:
        
    Next i
    
    
    '�e�����擾
    ret.Raw = merge_lines
    
    '���\�b�h��
    Dim wk As String: wk = GetMatchByRegExp(merge_lines, "(Function|Sub)\s+[A-Za-z_][A-Za-z0-9_]*\(", True)(0)
    ret.Name = Replace(Replace(wk, methodType & " ", ""), "(", "")
    
    start_clm = InStr(merge_lines, "(")
    end_clm = FindMatchingBracketPositionForVB(merge_lines)
    
    '�߂�l
    If methodType = "Function" Then
        wk = Mid(merge_lines, end_clm)
        
        If wk = ")" Then
            '�߂�l���Ȃ�Function
            ret.ret = ""
        Else
            ret.ret = Replace(wk, ") As ", "")
        End If
    Else
        ret.ret = ""
    End If
    
    '����
    wk = left(Mid(merge_lines, start_clm + 1), end_clm - start_clm - 1)
    ret.params = Split(wk, ",")
    
    For i = 0 To UBound(ret.params)
        ret.params(i) = Trim(ret.params(i))
    Next i
    
    GetMethodInfoForVB = ret

End Function

'-------------------------------------------------------------
' �ŏ��Ɍ�����"("�ɑΉ�����")"��T���Č��ʒu��Ԃ�
' inputString : I : �Ώە�����
' Ret : 0=������Ȃ�����, 1�ȏ�=")"�̌��ʒu
'-------------------------------------------------------------
Public Function FindMatchingBracketPositionForVB(ByVal inputString As String) As Long

    Dim ret As Long: ret = 0

    '�����`�F�b�N
    If inputString = "" Or _
       InStr(inputString, "(") = 0 Or InStr(inputString, ")") = 0 Then
        FindMatchingBracketPositionForVB = ret
        Exit Function
    End If

    Dim openBracket As String
    Dim closeBracket As String
    Dim i As Long
    Dim stack As Long
    
    openBracket = "("
    closeBracket = ")"
    
    ' �X�^�b�N��������
    stack = 0
    
    For i = 1 To Len(inputString)
        If Mid(inputString, i, 1) = openBracket Then
            ' �J�n����
            stack = stack + 1
        ElseIf Mid(inputString, i, 1) = closeBracket Then
            ' �I������
            stack = stack - 1
            If stack = 0 Then
                ' �Ή�����I�����ʂ���������
                FindMatchingBracketPositionForVB = i
                Exit Function
            End If
        End If
    Next i
    
    '�Ή�����I�����ʂ�������Ȃ�
    FindMatchingBracketPositionForVB = ret
    
End Function

'-------------------------------------------------------------
' A1�`���̕����񂩂��ԍ���Ԃ�
'  Ex. "A1" -> 1
'      "ZZ12" -> 27
' a1�FI : A1�`���̕�����
' Ret�F��ԍ�
'-------------------------------------------------------------
Public Function GetColNumFromA1(ByVal a1 As String) As Long
    Dim Clm As String
    Dim row As Long
    Call SplitCellAddress(a1, Clm, row)
    
    Dim substr As String
    substr = Clm
    
    Dim i As Integer
    
    For i = 1 To Len(Clm)
        If (i = 1) Then
            GetColNumFromA1 = GetColNumFromA1 + A_to_ColNum(right(substr, 1))
        Else
            GetColNumFromA1 = GetColNumFromA1 + (A_to_ColNum(right(substr, 1)) * (i - 1) * 26)
        End If
        substr = left(substr, Len(substr) - 1)
    Next i
End Function

'-------------------------------------------------------------
' A�`Z��1�`26�ɕϊ�����
' idx�FI : A�`Z
' Ret�F1�`26
'-------------------------------------------------------------
Public Function A_to_ColNum(ByVal az As String) As Integer
    A_to_ColNum = Asc(az) - 65 + 1
End Function

'-------------------------------------------------------------
' A1�`���̃Z���A�h���X��񖼂ƍs�ԍ��ɕ������ĕԂ�
' cell_adr : I : A1�`���̃Z���A�h���X
' clm_name : O : ��
' row_num : O : �s�ԍ�
' Ret : True=Success, False=Failed
'-------------------------------------------------------------
Public Function SplitCellAddress(ByVal cell_adr As String, ByRef clm_name As String, ByRef row_num As Long) As Boolean
    Dim matches() As String
    
    matches = GetMatchByRegExp(cell_adr, "[A-Z]+", True)
    
    If Common.IsEmptyArray(matches) = True Then
         SplitCellAddress = False
         Exit Function
    End If
    
    clm_name = matches(0)
    
    matches = GetMatchByRegExp(cell_adr, "[0-9]+", False)
    
    If Common.IsEmptyArray(matches) = True Then
         SplitCellAddress = False
         Exit Function
    End If
    
    row_num = CLng(matches(0))
    
    SplitCellAddress = True
End Function

'-------------------------------------------------------------
' �t�@�C���ɕ����񃊃X�g��UTF-8�ŏ�������
' path : I : �w��t�@�C���p�X(��΃p�X)
' str_ary : I : �����񃊃X�g
'-------------------------------------------------------------
Public Sub SaveToFileFromStringArray(ByVal path As String, ByRef str_ary() As String)
    If path = "" Or IsExistsFile(path) = False Then
        Err.Raise 53, , "[SaveToFileFromStringArray] �w�肳�ꂽ�p�X���s���ł� (path=" & path & ")"
    End If

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
On Error GoTo Error
    '�����R�[�h��UTF-8�ɐݒ肷��
    stream.Charset = "UTF-8"
    
    '�e�L�X�g���[�h�ŊJ��
    stream.Open
    
    Dim row As Long
    Dim line As String
    
    For row = 0 To UBound(str_ary)
        line = str_ary(row)
        stream.WriteText line
        stream.WriteText vbCrLf
    Next row
    
    Const OVER_WRITE = 2
    stream.SaveToFile path, OVER_WRITE

    stream.Close
    Set stream = Nothing
    
    Exit Sub
Error:
    stream.Close
    Set stream = Nothing
    
    Err.Raise 53, , "[SaveToFileFromStringArray] �G���[! (path=" & path & "), Desc=" & Err.Description
End Sub

'-------------------------------------------------------------
'������𖖔�����擪�Ɍ������Č��Ă����A�w�肳�ꂽ�������������炻���܂ł̕������Ԃ�
' ��:str="ABC:DEF", last_char=":"�̏ꍇ�A"DEF"���Ԃ�
' str : I : ������
' last_char : I : �w�肳�ꂽ����(1����)
' Ret : �w�肳�ꂽ�������������炻���܂ł̕�����B������Ȃ��ꍇ��""
'-------------------------------------------------------------
Public Function GetStringLastChar(ByVal str As String, ByVal last_char As String) As String
    '������̒������擾
    Dim length As Integer
    Dim i As Integer
    Dim ch As String
    
    length = Len(str)
    
    If length = 0 Then
        GetStringLastChar = ""
        Exit Function
    End If
    
    '������̖�������擪�Ɍ������ă��[�v
    For i = length To 1 Step -1
        'i�Ԗڂ̕������擾
        ch = Mid(str, i, 1)
        
        '��������
        If ch = last_char Then
            GetStringLastChar = right(str, length - i)
            Exit Function
        End If
    Next i
    
    '������Ȃ�����
    GetStringLastChar = ""
End Function

'-------------------------------------------------------------
'�p�X��255byte�ȏォ��Ԃ�
' path : I : �p�X (��΁E���΂̓`�F�b�N���Ȃ�)
' Ret : True/False (True=255byte�ȏ�, False=255byte����)
'-------------------------------------------------------------
Public Function IsMaxOverPath(ByVal path As String) As Boolean
    IsMaxOverPath = LenB(StrConv(path, vbFromUnicode)) >= 255
End Function

'-------------------------------------------------------------
'�����񂪎w�蕶����ŊJ�n����Ă��邩��Ԃ�
' target : I : ������
' search : I : �w�蕶����
' Ret : True/False (True=�J�n����Ă���, False=�J�n����Ă��Ȃ�)
'-------------------------------------------------------------
Public Function StartsWith(ByVal target As String, ByVal search As String) As Boolean
    StartsWith = False
    
    If Len(search) > Len(target) Then
        Exit Function
    End If
    
    If left(target, Len(search)) = search Then
        StartsWith = True
    End If
    
End Function

'-------------------------------------------------------------
'�t�H���_���󂩂ǂ�����Ԃ�
' path : I : �t�H���_�p�X(��΃p�X)
' Ret : True/False (True=��, False=��ł͖���)
'-------------------------------------------------------------
Public Function IsEmptyFolder(ByVal path As String) As Boolean
    If IsExistsFolder(path) = False Then
        Err.Raise 53, , "[IsEmptyFolder] �w�肳�ꂽ�t�H���_�����݂��܂��� (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsEmptyFolder] �p�X���������܂� (path=" & path & ")"
    End If

    Dim fso As Object
    Dim folder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(path)
    
    IsEmptyFolder = folder.Files.count = 0 And folder.SubFolders.count = 0
    
    Set fso = Nothing
    Set folder = Nothing
End Function

'-------------------------------------------------------------
'String�z��������\�[�g���ďd���s���폜���ĕԂ�
' arr : I : �z��
' Ret : �����\�[�g���ďd���s���폜�����z��
'-------------------------------------------------------------
Public Function SortAndDistinctArray(ByRef arr() As String) As String()
    Dim Dict As Object
    Set Dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Not Dict.Exists(arr(i)) Then
            Dict.Add arr(i), 1
        End If
    Next i
    Dim result() As String
    ReDim result(0 To Dict.count - 1)
    Dim key As Variant
    i = 0
    For Each key In Dict.Keys()
        result(i) = key
        i = i + 1
    Next key
    Set Dict = Nothing
    SortAndDistinctArray = result
End Function

'-------------------------------------------------------------
'�E�̃R�����g���폜���ĕԂ�
' str : I : ������
' ext : I : �g���q(Ex. "bas", "vb") ��VB�n�̂݃T�|�[�g
' Ret : �R�����g������΍폜���ĕԂ��B�Ȃ���Ό��̕������Ԃ�
' Ex. "abc 'def" �� "abc"
'-------------------------------------------------------------
Public Function RemoveRightComment(ByVal str As String, ByVal Ext As String) As String
    Dim pos As Long
    Dim ret As String
    
    If Ext = "bas" Or _
       Ext = "frm" Or _
       Ext = "cls" Or _
       Ext = "ctl" Or _
       Ext = "vb" Then
        pos = InStr(str, "'")
        
        If pos = 0 Then
            ret = str
        Else
            ret = RTrim(Mid(str, 1, pos - 1))
        End If
    Else
        Err.Raise 53, , "[RemoveRightComment] �w�肳�ꂽ�g���q�͖��T�|�[�g�ł� (ext=" & Ext & ")"
    End If
    
    RemoveRightComment = RTrim(ret)

End Function

'-------------------------------------------------------------
'�ŏ��Ɍ��������p���̈ʒu��Ԃ�
' str : I : ������
' Ret : �p���̈ʒu(������Ȃ��ꍇ��0��Ԃ�)
'-------------------------------------------------------------
Public Function FindFirstCasePosition(ByVal str As String) As Long
    Dim i As Long
    Dim char As String
    
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        If char Like "[A-Za-z]" Then
            FindFirstCasePosition = i
            Exit Function
        End If
    Next i
    
    FindFirstCasePosition = 0
End Function

'-------------------------------------------------------------
'�R�����g�s���𔻒肷��
' line : I : �s�f�[�^
' ext : I : �g���q(Ex. "bas", "vb") ��VB�n�̂݃T�|�[�g
' Ret : True/False(True=�R�����g�s)
'-------------------------------------------------------------
Public Function IsCommentCode(ByVal line As String, ByVal Ext As String) As Boolean
    If line = "" Or Ext = "" Then
        IsCommentCode = False
        Exit Function
    End If
    
    Dim wk As String
    wk = Replace(line, vbTab, " ")
    
    If Ext = "bas" Or _
       Ext = "frm" Or _
       Ext = "cls" Or _
       Ext = "ctl" Or _
       Ext = "vb" Then
        If left(LTrim(wk), 1) = "'" Or _
           left(LTrim(wk), 4) = "REM " Then
           IsCommentCode = True
           Exit Function
        End If
    Else
        Err.Raise 53, , "[IsCommentCode] �w�肳�ꂽ�g���q�͖��T�|�[�g�ł� (ext=" & Ext & ")"
    End If
    
    IsCommentCode = False

End Function

'-------------------------------------------------------------
'�t�H���_�p�X�Ɏw��t�H���_�������邩�`�F�b�N���A����΂��̃t�H���_�܂ł̃p�X��Ԃ�
' path : I : �t�H���_�p�X(��΃p�X)
' keyword : I : �L�[���[�h
' Ret : �L�[���[�h�܂ł̃p�X(�L�[���[�h��������Ȃ��ꍇ�͋��Ԃ�)
'-------------------------------------------------------------
Public Function GetFolderPathByKeyword( _
    path As String, _
    keyword As String _
) As String
    If path = "" Or keyword = "" Then
        GetFolderPathByKeyword = ""
        Exit Function
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[GetFolderPathByKeyword] �p�X���������܂� (path=" & path & ")"
    End If

    Dim SEP As String: SEP = Application.PathSeparator
    Dim path_ary() As String
    Dim ret_ary() As String
    Dim i As Integer
    Dim j As Integer
    
    path_ary = Split(path, SEP)

    For i = UBound(path_ary) To 0 Step -1
        If path_ary(i) = keyword Then
        
            ReDim Preserve ret_ary(i)
            
            For j = LBound(ret_ary) To UBound(ret_ary)
                ret_ary(j) = path_ary(j)
            Next j
        
            GetFolderPathByKeyword = Join(ret_ary, SEP)
            Exit Function
        End If
    Next i
    
    GetFolderPathByKeyword = ""
End Function

'-------------------------------------------------------------
'�t�H���_�p�X����Ō�̃t�H���_����Ԃ�
' path : I : �t�H���_�p�X(��΃p�X)
' Ret : �Ō�̃t�H���_��
'        ��: "C:\abc\def\xyz" �� "xyz"
'-------------------------------------------------------------
Public Function GetLastFolderName(ByVal path As String) As String
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[GetLastFolderName] �p�X���������܂� (path=" & path & ")"
    End If

    Dim new_path As String: new_path = Common.RemoveTrailingBackslash(path)
    GetLastFolderName = right(new_path, Len(new_path) - InStrRev(new_path, Application.PathSeparator))
End Function

'-------------------------------------------------------------
'�t�H���_�p�X�̖����Ɍ��ݓ����̕������t�^���ĕԂ�
' path : I : �t�H���_�p�X(��΃p�X)
' Ret : �����Ɍ��ݓ����̕������t�^�����t�@�C���p�X
'-------------------------------------------------------------
Public Function ChangeUniqueDirPath(ByVal path As String) As String
    If IsExistsFolder(path) = False Then
        Err.Raise 53, , "[ChangeUniqueDirPath] �w�肳�ꂽ�t�H���_�����݂��܂��� (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[ChangeUniqueDirPath] �p�X���������܂� (path=" & path & ")"
    End If

    Dim new_path As String
    new_path = path & "_" & GetNowTimeString()
    If IsExistsFolder(new_path) = True Then
        WaitSec 1
        new_path = path & "_" & GetNowTimeString()
    End If

    ChangeUniqueDirPath = new_path
End Function

'-------------------------------------------------------------
'���K�\���Ńp�^�[���}�b�`���O�����������u������
' test_str : I : �Ώە�����
' ptn : I : �����p�^�[��
' replace_str : I : �u���㕶����
' is_ignore_case : I : �啶������������ʂ��邩(True=����)
' Ret : �u����̑Ώە�����
' Note:
'  - �Q�Ɛݒ�Ɉȉ���ǉ�����
'    Microsoft VBScript Regular Expression 5.5
'-------------------------------------------------------------
Public Function ReplaceByRegExp( _
    ByVal test_str As String, _
    ByVal ptn As String, _
    ByVal replace_str As String, _
    ByVal is_ignore_case As Boolean _
) As String
    Dim REG As New VBScript_RegExp_55.RegExp
    Dim mc As MatchCollection
    Dim m As Match
    Dim list() As String
    ReDim list(0)
    
    REG.Global = True
    REG.ignoreCase = is_ignore_case
    REG.Pattern = ptn
    
    ReplaceByRegExp = REG.Replace(test_str, replace_str)

End Function

'-------------------------------------------------------------
'���K�\���Ńp�^�[���}�b�`���O�������ʂ�Ԃ�
' test_str : I : �Ώە�����
' ptn : I : �����p�^�[��
' is_ignore_case : I : �啶������������ʂ��邩(True=����)
' Ret : �}�b�`���������񃊃X�g
' Note:
'  - �Q�Ɛݒ�Ɉȉ���ǉ�����
'    Microsoft VBScript Regular Expression 5.5
'-------------------------------------------------------------
Public Function GetMatchByRegExp( _
    ByVal test_str As String, _
    ByVal ptn As String, _
    ByVal is_ignore_case As Boolean _
) As String()
    Dim REG As New VBScript_RegExp_55.RegExp
    Dim mc As MatchCollection
    Dim m As Match
    Dim list() As String
    ReDim list(0)
    
    REG.Global = True
    REG.ignoreCase = is_ignore_case
    REG.Pattern = ptn
    
    Set mc = REG.Execute(test_str)
    For Each m In mc
        Common.AppendArray list, m.value
    Next
    
    list = Common.DeleteEmptyArray(list)
    
    GetMatchByRegExp = list
End Function

'-------------------------------------------------------------
'���K�\���Ńp�^�[���}�b�`���O���s��
' test_str : I : �Ώە�����
' ptn : I : �����p�^�[��
' is_ignore_case : I : �啶������������ʂ��邩(True=����)
' Ret : True/False (True=��v)
' Note:
'  - �Q�Ɛݒ�Ɉȉ���ǉ�����
'    Microsoft VBScript Regular Expression 5.5
'-------------------------------------------------------------
Public Function IsMatchByRegExp( _
    ByVal test_str As String, _
    ByVal ptn As String, _
    ByVal is_ignore_case As Boolean _
) As Boolean
    Dim REG As New VBScript_RegExp_55.RegExp
    REG.Global = True
    REG.ignoreCase = is_ignore_case
    REG.Pattern = ptn
    
    IsMatchByRegExp = REG.Test(test_str)
End Function

'-------------------------------------------------------------
'���g�̃t�H���_�p�X��Ԃ�
' Ret : �t�H���_�p�X
'-------------------------------------------------------------
Public Function GetMyDir() As String
    Dim currentProject As Workbook
    Set currentProject = ThisWorkbook
    GetMyDir = currentProject.path
End Function

'-------------------------------------------------------------
'������z���A�����ĕ������Ԃ�
' ary : I : ������z��
' delim : I : ��؂蕶��(1����)
' with_dbl_quot : I : �_�u���N�H�[�e�[�V�����ň͂ނ��ۂ� (True=�͂�)
' Ret : ��؂蕶���ŘA����̕�����
'-------------------------------------------------------------
Public Function JoinFromArray(ByRef ary() As String, ByVal delim As String, ByVal with_dbl_quot As Boolean) As String
    If IsEmptyArray(ary) = True Or delim = "" Then
        JoinFromArray = ""
        Exit Function
    End If

    Dim ret As String: ret = ""
    Dim i As Long
    
    For i = LBound(ary) To UBound(ary)
        If with_dbl_quot = True Then
            ret = ret & Chr(34) & ary(i) & Chr(34) & delim
        Else
            ret = ret & ary(i) & delim
        End If
    Next i
    
    JoinFromArray = left(ret, Len(ret) - 1)

End Function

'-------------------------------------------------------------
'�u�b�N���J���Ă��邩�ۂ���Ԃ�
' book_name : I : �u�b�N��
' Ret : True/False (True=�J���Ă���)
'-------------------------------------------------------------
Function IsOpenWorkbook(ByVal book_name As String) As Boolean
    Dim wb As Workbook
    Dim is_err As Boolean
    is_err = False

On Error Resume Next
    Set wb = Workbooks(book_name)
    
    If Err.Number <> 0 Then
        is_err = True
        Err.Clear
    End If

On Error GoTo 0
    If is_err = True Then
        IsOpenWorkbook = False
    Else
        IsOpenWorkbook = True
    End If
End Function

'-------------------------------------------------------------
'��t�@�C�����ۂ���Ԃ�
' path : I : �t�@�C���p�X(��΃p�X)
' Ret : True/False (True=��t�@�C��)
'-------------------------------------------------------------
Public Function IsEmptyFile(ByVal path As String) As Boolean
    If IsExistsFile(path) = False Then
        Err.Raise 53, , "[IsEmptyFile] �w�肳�ꂽ�t�@�C�������݂��܂��� (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsEmptyFile] �p�X���������܂� (path=" & path & ")"
    End If

    IsEmptyFile = (FileLen(path) = 0)
End Function

'-------------------------------------------------------------
'Variant�^�̔z���String�^�̔z��ɕϊ�����
' arr : I : variant�^�̔z��
' Ret : String�^�̔z��
'-------------------------------------------------------------
Public Function VariantToStringArray(arr As Variant) As String()
    Dim ret_arr() As String
    Dim i As Long
    
    ReDim ret_arr(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        ret_arr(i) = CStr(arr(i))
    Next i
    
    VariantToStringArray = ret_arr
End Function

'-------------------------------------------------------------
'�t�@�C�����̃L�[���[�h���܂ލs���폜���ď㏑���ۑ�����
' path : I : �t�@�C���p�X(��΃p�X)
' keyword : I : �L�[���[�h
'-------------------------------------------------------------
Public Sub RemoveLinesWithKeyword(ByVal path As String, ByVal keyword As String)
    If IsExistsFile(path) = False Then
        Err.Raise 53, , "[RemoveLinesWithKeyword] �w�肳�ꂽ�t�@�C�������݂��܂��� (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[RemoveLinesWithKeyword] �p�X���������܂� (path=" & path & ")"
    End If

    If keyword = "" Then
        Exit Sub
    End If
    
    Dim fso As Object
    Dim file As Object
    Dim temp_file As Object
    Dim line As String
    Dim temp_ext As String: temp_ext = "." & GetNowTimeString()
    
    Const READ_ONLY = 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(path, READ_ONLY)
    Set temp_file = fso.CreateTextFile(path & temp_ext, True)
    
    Do While Not file.AtEndOfStream
        line = file.ReadLine
        
        If InStr(line, keyword) = 0 Then
            temp_file.WriteLine line
        End If
    Loop
    
    file.Close
    temp_file.Close
    
    fso.DeleteFile path
    fso.MoveFile path & temp_ext, path
    
    Set temp_file = Nothing
    Set file = Nothing
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'�����񂩂�L�[���[�h�Ō������A�q�b�g�����L�[���[�h����Ō�܂ł̕������Ԃ�
' target : I : �����Ώۂ̕�����
' keyword : I : �����L�[���[�h
' Ret : �q�b�g�����L�[���[�h����Ō�܂ł̕�����(������Ȃ��ꍇ��"")
' Ex.
'   target:"C:\abc\def\xyz\123.txt"
'   keyword:"def"
'   Ret:"def\xyz\123.txt"
'-------------------------------------------------------------
Function GetStringByKeyword(ByVal target As String, ByVal keyword As String) As String
    Dim pos As Long
    pos = InStr(target, keyword)
    If pos > 0 Then
        GetStringByKeyword = Mid(target, pos)
    Else
        GetStringByKeyword = ""
    End If
End Function

'-------------------------------------------------------------
'Git�R�}���h�����s����
' repo_path : I : ���[�J�����|�W�g���t�H���_�p�X(��΃p�X)
' command : I : �R�}���h (Ex."git log --oneline")
' Ret : �W���o��
'-------------------------------------------------------------
Public Function RunGit(ByVal repo_path As String, ByVal command As String) As String()
    Dim err_msg As String: err_msg = ""
    Dim std_out() As String

    If IsMaxOverPath(repo_path) = True Then
        Err.Raise 53, , "[RunGit] �p�X���������܂� (repo_path=" & repo_path & ")"
    End If

    If IsExistsFile(GIT_BASH) = False Then
        err_msg = "[RunGit] git��������܂��� (" & GIT_BASH & ")"
        GoTo FINISH_3
    End If
    
    If IsExistsFolder(repo_path) = False Then
        If InStr(command, "git clone") = 0 Then
            err_msg = "[RunGit] �w�肳�ꂽ�t�H���_�����݂��܂��� (repo_path=" & repo_path & ")"
            GoTo FINISH_3
        End If
    End If
    
    '�R�}���h���s���ʊi�[�p�̈ꎞ�t�@�C���p�X
    Dim temp As String: temp = GetTempFolder() & Application.PathSeparator & GetNowTimeString() & ".txt"

    '�R�}���h�쐬
    Dim run_cmd As String: run_cmd = GIT_BASH & _
                                     " --login -i -c & cd " & repo_path & " & " & _
                                     command & _
                                     " > " & temp & " 2>&1"
    WriteLog "[RunGit] run_cmd=" & run_cmd
    
    '�R�}���h���s
    Dim objShell As Object
    Dim objExec As Object
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("cmd.exe /c " & Chr(34) & run_cmd & Chr(34))
    
    '�v���Z�X�������ɒʒm���󂯎��
    Do While objExec.Status = 0
        DoEvents
    Loop
    
    '�v���Z�X�̖߂�l���擾����
    If objExec.ExitCode <> 0 Then
        err_msg = "[RunGit] �v���Z�X�̖߂�l��0�ȊO�ł� (exit code=" & objExec.ExitCode & ")"
        
        If IsEmptyFile(temp) = True Then
            GoTo FINISH_2
        Else
            GoTo FINISH
        End If
        
    End If
    
    If IsEmptyFile(temp) = True Then
        GoTo FINISH_2
    End If
    
FINISH:
    If IsUTF8(temp) = False Then
        std_out = Split(ReadTextFileBySJIS(temp), vbCrLf)
    Else
        'std_out = Split(ReadTextFileByUTF8(temp), vbLf)
        std_out = Split(Replace(ReadTextFileByUTF8(temp), vbCrLf, vbLf), vbLf)
    End If

FINISH_2:
    DeleteFile (temp)
    
FINISH_3:
    Set objShell = Nothing
    Set objExec = Nothing
    
    If err_msg <> "" Then
        Err.Raise 53, , err_msg & vbCrLf & "std_out=" & Join(std_out, ",")
    End If

    RunGit = std_out
End Function

'-------------------------------------------------------------
'�ꎞ�t�H���_�p�X���擾����
' Ret : �ꎞ�t�H���_�p�X(��΃p�X)
'-------------------------------------------------------------
Public Function GetTempFolder() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetTempFolder = fso.getSpecialFolder(2)
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'�t�@�C�����R�s�[����
' src_path : I : �R�s�[���t�@�C���p�X(��΃p�X)
' dst_path : I : �R�s�[��t�@�C���p�X(��΃p�X)
' is_create_dir : I : �R�s�[��Ƀt�H���_���쐬���ăR�s�[����(�f�t�H���g=false)
'-------------------------------------------------------------
Public Sub CopyFile(ByVal src_path As String, ByVal dst_path As String, Optional ByVal is_create_dir As Boolean = False)
    If IsExistsFile(src_path) = False Then
        Err.Raise 53, , "[CopyFile] �w�肳�ꂽ�t�@�C�������݂��܂��� (src_path=" & src_path & ")"
    End If

    If IsMaxOverPath(src_path) = True Or IsMaxOverPath(dst_path) = True Then
        Err.Raise 53, , "[CopyFile] �p�X���������܂� (src_path=" & src_path & ", dst_path=" & dst_path & ")"
    End If

    If dst_path = "" Or src_path = dst_path Or IsExistsFile(dst_path) = True Then
        Exit Sub
    End If
    
    If is_create_dir = False Then
        FileCopy src_path, dst_path
        Exit Sub
    End If
    
    Dim dst_dir_path As String
    dst_dir_path = Common.GetFolderPath(dst_path)
    
    If Common.IsExistsFolder(dst_dir_path) = True Then
        FileCopy src_path, dst_path
        Exit Sub
    End If
    
    '�R�s�[��t�H���_�p�X�����݂��Ȃ��̂ō쐬����
    Common.CreateFolder (dst_dir_path)
    
    FileCopy src_path, dst_path
End Sub


'-------------------------------------------------------------
'�t�H���_�����l�[������
' path : I : �t�H���_�p�X(��΃p�X)
' rename : I : ���l�[����̃t�H���_��
' Ret : ���l�[����̃t�H���_�p�X
'-------------------------------------------------------------
Public Function RenameFolder(ByVal path As String, ByVal rename As String) As String
    If IsExistsFolder(path) = False Then
        Err.Raise 53, , "[RenameFolder] �w�肳�ꂽ�t�H���_�����݂��܂��� (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[RenameFolder] �p�X���������܂� (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(path)
    
    Dim err_msg As String
    Dim retry As Integer
    For retry = 0 To 3

On Error Resume Next
        folder.Name = rename
    
        err_msg = Err.Description
        Err.Clear
On Error GoTo 0

        If err_msg = "" Then
            Exit For
        End If
        
        WaitSec 1

    Next retry
    
    Set fso = Nothing
    
    If err_msg <> "" Then
        Err.Raise 53, , "[RenameFolder] �G���[! (err_msg=" & err_msg & ")"
    End If

    RenameFolder = folder.path

End Function

'-------------------------------------------------------------
'���[�N�V�[�g�̎w���̃f�[�^�ŏI�s�ԍ���Ԃ�
' ws : I : ���[�N�V�[�g
' clm : I : �w���(Ex."A")
'-------------------------------------------------------------
Public Function GetLastRowFromWorksheet( _
  ByVal ws As Worksheet, _
  ByVal Clm As String _
) As Long
    GetLastRowFromWorksheet = ws.Cells(ws.Rows.count, Clm).End(xlUp).row
End Function

'-------------------------------------------------------------
'������̔z�񂩂�w�胏�[�h�Ō������A�q�b�g�����s�ԍ���Ԃ�
' keyword : I : �������[�h
' input_array : I : ������̔z��
' is_use_regexp : I : ���K�\���̎g�p�L��
' Ret : �q�b�g�����s�ԍ�
'-------------------------------------------------------------
Public Function FindRowByKeywordFromArray(ByVal keyword As String, ByRef input_array() As String, ByVal is_use_regexp As Boolean) As Long
    If keyword = "" Then
        FindRowByKeywordFromArray = -1
        Exit Function
    End If

    Dim row As Long
    Dim isMatch As Boolean
    Dim line As String
    Dim regex As Object
    Set regex = Nothing
    
    If is_use_regexp = True Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = keyword
    End If
   
    For row = LBound(input_array) To UBound(input_array)
        line = input_array(row)
        
        If is_use_regexp = True Then
            isMatch = regex.Test(line)
        ElseIf InStr(1, line, keyword) > 0 Then
            isMatch = True
        End If
    
        If isMatch = True Then
            FindRowByKeywordFromArray = row
            Exit Function
        End If
    Next row
    
    FindRowByKeywordFromArray = -1
End Function

'-------------------------------------------------------------
'���[�N�V�[�g�̎w���̑S�s���w�胏�[�h�Ō������A�q�b�g�����s�ԍ����X�g��Ԃ�
' ws : I : ���[�N�V�[�g
' find_clm : I : �w���(Ex."A")
' find_start_row : I : �����J�n�s(1�n�܂�)
' find_end_row : I : �����I���s(0�̏ꍇ�͑S�s�Ƃ���)
' keyword : I : �������[�h
' Ret : �q�b�g�����s�ԍ����X�g
'-------------------------------------------------------------
Public Function FindRowListByKeywordFromWorksheet( _
  ByVal ws As Worksheet, _
  ByVal find_clm As String, _
  ByVal find_start_row As Long, _
  ByVal find_end_row As Long, _
  ByVal keyword As String _
) As Long()
    Dim rng As Range
    Dim cell As Range
    Dim cnt As Long
    Dim found_rows() As Long
    
    If find_start_row > find_end_row Or find_end_row < 0 Then
        find_end_row = 0
    End If
    
    If find_end_row = 0 Then
        Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(ws.Rows.count, find_clm).End(xlUp).row)
    Else
        Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(find_end_row, find_clm).End(xlUp).row)
    End If

    cnt = 0
    
    For Each cell In rng
        If cell.value = keyword Then
            ReDim Preserve found_rows(cnt)
            found_rows(cnt) = cell.row
            cnt = cnt + 1
        End If
    Next cell
    
    FindRowListByKeywordFromWorksheet = found_rows
End Function

'-------------------------------------------------------------
'���[�N�V�[�g�̎w���̑S�s���w�胏�[�h�Ō������A�q�b�g�����s�ԍ���Ԃ�
' ws : I : ���[�N�V�[�g
' find_clm : I : �w���(Ex."A")
' find_start_row : I : �����J�n�s(1�n�܂�)
' keyword : I : �������[�h
' find_end_row : I : �����I���s(�C�ӁB0�̏ꍇ�͑S�s�Ƃ���)
' Ret : �q�b�g�����s�ԍ�
'-------------------------------------------------------------
Public Function FindRowByKeywordFromWorksheet( _
  ByVal ws As Worksheet, _
  ByVal find_clm As String, _
  ByVal find_start_row As Long, _
  ByVal keyword As String, _
  Optional ByVal find_end_row As Long = 0 _
) As Long
    Dim rng As Range
    Dim cell As Range
    Dim found_row As Long
    
    If find_start_row > find_end_row Or find_end_row < 0 Then
        find_end_row = 0
    End If
    
    If find_end_row = 0 Then
        Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(ws.Rows.count, find_clm).End(xlUp).row)
        'Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(ws.Rows.count, find_clm).End(xlDown).row)
    Else
        'Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(find_end_row, find_clm).End(xlUp).row)
        Set rng = ws.Range(find_clm & find_start_row & ":" & find_clm & ws.Cells(find_end_row, find_clm).End(xlDown).row)
    End If

    found_row = 0
    For Each cell In rng
        If cell.value = keyword Then
            found_row = cell.row
            Exit For
        End If
    Next cell
    
    FindRowByKeywordFromWorksheet = found_row
End Function

'-------------------------------------------------------------
'�V�[�g�̓��e��2�����z��Ɋi�[����
' sheet_name : I : �V�[�g��
' Ret : �V�[�g�̓��e
'-------------------------------------------------------------
Public Function GetSheetContentsByStringArray(ByVal sheet_name As String) As String()
    Dim ws As Worksheet
    Dim arr() As String
    Dim row_cnt As Long, clm_cnt As Long
    Dim r As Long, c As Long
    
    Set ws = ActiveWorkbook.Worksheets(sheet_name)

    row_cnt = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    clm_cnt = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ReDim arr(1 To row_cnt, 1 To clm_cnt)
    For r = 1 To row_cnt
        For c = 1 To clm_cnt
            arr(r, c) = CStr(ws.Cells(r, c).value)
        Next c
    Next r

    GetSheetContentsByStringArray = arr
End Function

'-------------------------------------------------------------
'�g���q��ύX����
' path : I : �t�@�C���p�X(��΃p�X)
' ext : I : �ύX��̊g���q(Ex. ".new")
' Ret : �ύX��̃t�@�C���p�X(��΃p�X)
'       path�̃t�@�C�������݂��Ȃ��ꍇ��path��Ԃ�
'-------------------------------------------------------------
Public Function ChangeFileExt(ByVal path As String, ByVal Ext As String) As String
    If IsExistsFile(path) = False Then
        'Err.Raise 53, , "[ChangeFileExt] �w�肳�ꂽ�t�@�C�������݂��܂��� (path=" & path & ")"
        ChangeFileExt = path
        Exit Function
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[ChangeFileExt] �p�X���������܂� (path=" & path & ")"
    End If

    Dim SEP As String: SEP = Application.PathSeparator
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim old_ext As String: old_ext = fso.GetExtensionName(path)
    Dim file_name As String: file_name = fso.GetBaseName(path)
    Dim new_path As String
    
    '�V�����g���q�ɕύX
    file_name = file_name & Ext
    new_path = fso.GetParentFolderName(path) & SEP & file_name
    
    '�t�@�C������ύX
    fso.MoveFile path, new_path
    Set fso = Nothing
    
    ChangeFileExt = new_path
End Function

'-------------------------------------------------------------
'�u�b�N���J���ăV�[�g���擾����
' book_path : I : Excel�t�@�C���p�X(��΃p�X)
' sheet_name : I : �V�[�g��
' readonly : I : True/False (True=�ǎ��p�ŊJ��, False=�ǎ��p�ŊJ���Ȃ�)
' visible : I : True/False (True=�\��, False=��\��)
' Ret : �V�[�g�I�u�W�F�N�g
'-------------------------------------------------------------
Public Function GetSheet( _
    ByVal book_path As String, _
    ByVal sheet_name As String, _
    ByVal is_readonly As Boolean, _
    ByVal is_visible As Boolean _
) As Worksheet

    If IsMaxOverPath(book_path) = True Then
        Err.Raise 53, , "[GetSheet] �p�X���������܂� (book_path=" & book_path & ")"
    End If

    Dim wb As Workbook
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    
    If IsOpenWorkbook(book_path) = True Then
        '���ɊJ���Ă���
        Set wb = Workbooks(book_path)
    Else
        Set wb = Workbooks.Open(filename:=book_path, UpdateLinks:=False, ReadOnly:=is_readonly)
    End If
    
    wb.Activate
    ActiveWindow.Visible = is_visible
    
    If Common.IsExistSheet(wb, sheet_name) = False Then
        Err.Raise 53, , "[GetSheet] �w�肳�ꂽ�V�[�g�����݂��܂��� (book_path=" & book_path & ", sheet_name=" & sheet_name & ")"
    End If
    
    Set GetSheet = wb.Worksheets(sheet_name)

End Function

'-------------------------------------------------------------
'�u�b�N��ۑ����ĕ���
' name : I : �u�b�N��(Excel�t�@�C����)
'-------------------------------------------------------------
Public Sub SaveAndCloseBook(ByVal Name As String)
    Dim wb As Workbook
    For Each wb In Workbooks
        If InStr(wb.Name, Name) > 0 Then
            wb.Save
            wb.Close
        End If
    Next
End Sub

'-------------------------------------------------------------
'�u�b�N�����
' name : I : �u�b�N��(Excel�t�@�C����)
'-------------------------------------------------------------
Public Sub CloseBook(ByVal Name As String)
    Dim wb As Workbook
    For Each wb In Workbooks
        If InStr(wb.Name, Name) > 0 Then
            wb.Close SaveChanges:=False
        End If
    Next
End Sub

'-------------------------------------------------------------
'�t�@�C�����폜����
' path : IN : �t�@�C���p�X(��΃p�X)
'-------------------------------------------------------------
Public Sub DeleteFile(ByVal path As String)
    If IsExistsFile(path) = False Then
        Exit Sub
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[DeleteFile] �p�X���������܂� (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Const DELETE_READONLY = True
    fso.DeleteFile path, DELETE_READONLY
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'�t�@�C���������j�[�N�Ȗ��̂Ƀ��l�[�����ăR�s�[����
' src_file_path : IN : �R�s�[���t�@�C���p�X(��΃p�X)
' dst_dir_path : IN : �R�s�[��t�H���_�p�X(��΃p�X)
'                     ������\�͕s�v
'                     ��̏ꍇ�̓R�s�[���Ɠ����t�H���_�Ƃ���
' Ret : ���l�[���R�s�[��̃t�@�C���p�X
'-------------------------------------------------------------
Public Function CopyUniqueFile(ByVal src_file_path As String, ByVal dst_dir_path As String) As String
    If IsExistsFile(src_file_path) = False Then
        CopyUniqueFile = ""
        Exit Function
    End If

    If IsMaxOverPath(src_file_path) = True Or IsMaxOverPath(dst_dir_path) = True Then
        Err.Raise 53, , "[CopyUniqueFile] �p�X���������܂� (src_file_path=" & src_file_path & ", dst_dir_path=" & dst_dir_path & ")"
    End If

    Dim SEP As String: SEP = Application.PathSeparator
    Dim dst_file_path As String
    
    Dim unique_filename As String: unique_filename = GetFileName(src_file_path) & ".bak_" & GetNowTimeString()
    
    If dst_dir_path = "" Then
        dst_file_path = GetFolderNameFromPath(src_file_path) & SEP & unique_filename
    Else
        dst_file_path = dst_dir_path & SEP & unique_filename
    End If

    FileCopy src_file_path, dst_file_path
    
    CopyUniqueFile = dst_file_path
End Function

'-------------------------------------------------------------
'�t�@�C������Ԃ�
' path : IN : �t�@�C���p�X(��΃p�X)
' Ret : �t�@�C����
'-------------------------------------------------------------
Public Function GetFileName(ByVal path As String) As String
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[GetFileName] �p�X���������܂� (path=" & path & ")"
    End If
    
    If path = "" Then
        Err.Raise 53, , "[GetFileName] �p�X����ł� (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetFileName(path)
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'�t�@�C���p�X����t�H���_�p�X��Ԃ�
' path : IN : �t�@�C���p�X(��΃p�X)
' Ret : �t�H���_�p�X(��΃p�X)
'-------------------------------------------------------------
Public Function GetFolderPath(ByVal file_path As String) As String
    If IsMaxOverPath(file_path) = True Then
        Err.Raise 53, , "[GetFolderPath] �p�X���������܂� (file_path=" & file_path & ")"
    End If
    
    If file_path = "" Then
        Err.Raise 53, , "[GetFolderPath] �p�X����ł� (file_path=" & file_path & ")"
    End If
    
    Dim pos As Integer
    pos = InStrRev(file_path, "\")
    
    '\��荶���̕�������t�H���_�p�X�Ƃ��ĕԂ�
    GetFolderPath = left(file_path, pos - 1)
End Function

'-------------------------------------------------------------
'�w��t�H���_�z�����w��t�@�C�����Ō������ăt�@�C���p�X��Ԃ�
' search_path : IN : �����t�H���_�p�X(��΃p�X)
' search_name : IN : �����t�@�C����
' Ret : �t�@�C���p�X
'-------------------------------------------------------------
Public Function SearchFile(ByVal search_path As String, ByVal search_name As String) As String
    If IsMaxOverPath(search_path) = True Then
        Err.Raise 53, , "[SearchFile] �p�X���������܂� (search_path=" & search_path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(search_path)
    
    Dim file As Object
    For Each file In folder.Files
        If fso.FileExists(file.path) And fso.GetFileName(file.path) Like search_name Then
            '����
            SearchFile = file.path
            Set fso = Nothing
            Exit Function
        End If
    Next file
    
    '�T�u�t�H���_����������
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        Dim result As String
        result = SearchFile(subfolder.path, search_name)
        If result <> "" Then
            '�T�u�t�H���_���猋�ʂ��Ԃ��Ă����ꍇ�́A���̌��ʂ�Ԃ�
            SearchFile = result
            Set fso = Nothing
            Exit Function
        End If
    Next subfolder
    
    '�����Ώۂ̃t�@�C����������Ȃ������ꍇ
    SearchFile = ""
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'�w��t�H���_��UTF8��S��SJIS�ɂ���
' path : IN : �t�H���_�p�X(��΃p�X)
' ext : IN : �g���q(Ex."*.vb")
' is_subdir : IN : �T�u�t�H���_�܂ނ� (True=�܂�)
' Ret : �t�@�C�����X�g
'-------------------------------------------------------------
Public Sub UTF8toSJIS_AllFile(ByVal path As String, ByVal Ext As String, ByVal is_subdir As Boolean)
    If IsExistsFolder(path) = False Then
        Err.Raise 53, , "[UTF8toSJIS_AllFile] �w�肳�ꂽ�t�H���_�����݂��܂��� (path=" & path & ")"
    End If
    
    If Ext = "" Then
        Err.Raise 53, , "[UTF8toSJIS_AllFile] �g���q���w�肳��Ă��܂���"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[UTF8toSJIS_AllFile] �p�X���������܂� (path=" & path & ")"
    End If

    Dim i As Long
    Dim src_file_list() As String: src_file_list = CreateFileList(path, Ext, is_subdir)

    For i = LBound(src_file_list) To UBound(src_file_list)
        UTF8toSJIS src_file_list(i), False
    Next i
End Sub

'-------------------------------------------------------------
'�w��t�H���_��SJIS��S��UTF8�ɂ���
' path : IN : �t�H���_�p�X(��΃p�X)
' ext : IN : �g���q(Ex."*.vb")
' is_subdir : IN : �T�u�t�H���_�܂ނ� (True=�܂�)
' Ret : �t�@�C�����X�g
'-------------------------------------------------------------
Public Sub SJIStoUTF8_AllFile(ByVal path As String, ByVal Ext As String, ByVal is_subdir As Boolean)
    If IsExistsFolder(path) = False Then
        Err.Raise 53, , "[SJIStoUTF8_AllFile] �w�肳�ꂽ�t�H���_�����݂��܂��� (path=" & path & ")"
    End If
    
    If Ext = "" Then
        Err.Raise 53, , "[SJIStoUTF8_AllFile] �g���q���w�肳��Ă��܂���"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[SJIStoUTF8_AllFile] �p�X���������܂� (path=" & path & ")"
    End If

    Dim i As Long
    Dim src_file_list() As String: src_file_list = CreateFileList(path, Ext, is_subdir)

    For i = LBound(src_file_list) To UBound(src_file_list)
        SJIStoUTF8 src_file_list(i), False
    Next i
End Sub

'-------------------------------------------------------------
'�w�肳�ꂽ�t�@�C����SJIS��UTF8(BOM����)�ϊ�����
' path : IN : �t�@�C���p�X(��΃p�X)
' is_backup : IN : True/False (True=�o�b�N�A�b�v����)
'                  ��������".bak_���ݓ���"��t�^
'-------------------------------------------------------------
Public Sub SJIStoUTF8(ByVal path As String, ByVal is_backup As Boolean)
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[SJIStoUTF8] �p�X���������܂� (path=" & path & ")"
    End If

    Dim in_str As String
    Dim buf As String
    Dim i As Long
    
    Dim filenum As Integer: filenum = FreeFile
    
    'Shift-JIS�`���̃e�L�X�g�t�@�C����ǂݍ���
    in_str = ""
    Open path For Input As #filenum
        '�e�L�X�g�����ׂĎ擾����
        Do Until EOF(filenum)
            Line Input #filenum, buf
            in_str = in_str & buf & vbCrLf
        Loop
    Close #filenum
        
    'Shift-JIS�ȊO�̃t�@�C����ǂݍ���ł��܂����ꍇ�͏I��
    For i = 1 To Len(in_str)
        If Asc(Mid(in_str, i, 1)) = -7295 Then Exit Sub
    Next
    
    '�o�b�N�A�b�v
    If is_backup = True Then
        FileCopy path, path & ".bak_" & GetNowTimeString()
    End If
    
    'UTF-8�iBOM�t���j�Ńe�L�X�g�t�@�C���֏o��
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .WriteText in_str, 0
        .SaveToFile path, 2
        .Close
    End With
    
End Sub

'-------------------------------------------------------------
'�w�肳�ꂽ�t�@�C����UTF8(BOM����/�Ȃ�) �� SJIS�ϊ�����
' path : IN : �t�@�C���p�X(��΃p�X)
' is_backup : IN : True/False (True=�o�b�N�A�b�v����)
'                  ��������".bak_���ݓ���"��t�^
'-------------------------------------------------------------
Public Sub UTF8toSJIS(ByVal path As String, ByVal is_backup As Boolean)
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[UTF8toSJIS] �p�X���������܂� (path=" & path & ")"
    End If

    Dim in_str As String
    Dim out_str() As String
    Dim i As Long
    
    'UTF-8��������UTF-8�iBOM�t���j�̃e�L�X�g�t�@�C����ǂݍ���
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile path
        in_str = .ReadText
        .Close
    End With
    
    'UTF-8��������UTF-8�iBOM�t���j�ȊO��ǂݍ���ł��܂����ꍇ�͏I��
    For i = 1 To Len(in_str)
        If Mid(in_str, i, 1) <> Chr(63) Then
            If Asc(Mid(in_str, i, 1)) = 63 Then
                Exit Sub
            End If
        End If
    Next
    
    '���s���Ƀf�[�^�𕪂���
    out_str = Split(in_str, vbCrLf)
    
    '�o�b�N�A�b�v
    If is_backup = True Then
        FileCopy path, path & ".bak_" & GetNowTimeString()
    End If
    
    Dim filenum As Integer: filenum = FreeFile
    
    'Shift-JIS�`���Ńe�L�X�g�t�@�C���֏o��
    Open path For Output As #filenum
        For i = 0 To UBound(out_str)
            Print #filenum, out_str(i)
        Next
    Close #filenum

End Sub

'-------------------------------------------------------------
'�t�@�C����SJIS���𔻒肷��
' path : IN : �t�@�C���p�X(��΃p�X)
' Ret : True/False (True=SJIS)
'-------------------------------------------------------------
Public Function IsSJIS(ByVal path As String) As Boolean
    If IsExistsFile(path) = False Then
        Err.Raise 53, , "[IsSJIS] �w�肳�ꂽ�t�@�C�������݂��܂��� (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsSJIS] �p�X���������܂� (path=" & path & ")"
    End If

    Dim Ado As Object
    Const TYPE_BINARY = 1
    Set Ado = CreateObject("ADODB.Stream")
    Ado.Type = TYPE_BINARY
    Ado.Open

    Ado.LoadFromFile path
    Dim read_data As String: read_data = Ado.Read
    Ado.Close
    Set Ado = Nothing

    Dim i As Long
    Dim first_byte As Byte
    Dim second_byte As Byte
    Dim is_dbcs As Boolean
    
    For i = 1 To LenB(read_data)

        first_byte = AscB(MidB(read_data, i, 1))

        '�S�p������(DBCS)�̐擪1�o�C�g�ł��邩
        is_dbcs = False

        If &H81 <= first_byte And first_byte <= &H9F Then
            is_dbcs = True
        ElseIf &HE0 <= first_byte And first_byte <= &HEF Then
            is_dbcs = True
        End If

        If is_dbcs Then
            i = i + 1

            If i > LenB(read_data) Then
                IsSJIS = False
                Exit Function
            End If

            second_byte = AscB(MidB(read_data, i, 1))

            If &H40 <= second_byte And second_byte <= &H7F Then
                'SJIS!
            ElseIf &H80 <= second_byte And second_byte <= &HFC Then
                'SJIS!
            Else
                IsSJIS = False
                Exit Function
            End If
        End If
    Next

    IsSJIS = True
End Function

'-------------------------------------------------------------
'�t�@�C����UTF8(BOM����/�Ȃ�)���𔻒肷��
' path : IN : �t�@�C���p�X(��΃p�X)
' Ret : True/False (True=UTF8(BOM����/�Ȃ�))
'-------------------------------------------------------------
Public Function IsUTF8(ByVal path As String) As Boolean
    If IsExistsFile(path) = False Then
        Err.Raise 53, , "[IsUTF8] �w�肳�ꂽ�t�@�C�������݂��܂��� (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsUTF8] �p�X���������܂� (path=" & path & ")"
    End If

    Dim in_str As String
    Dim out_str() As String
    Dim i As Long
    
    'UTF-8��������UTF-8�iBOM�t���j�̃e�L�X�g�t�@�C����ǂݍ���
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile path
        in_str = .ReadText
        .Close
    End With
    
    'UTF-8��������UTF-8�iBOM�t���j�ȊO��ǂݍ���ł��܂����ꍇ�͏I��
    For i = 1 To Len(in_str)
        If Mid(in_str, i, 1) <> Chr(63) Then
            If Asc(Mid(in_str, i, 1)) = 63 Then
                IsUTF8 = False
                Exit Function
            End If
        End If
    Next
    
    IsUTF8 = True
End Function

'-------------------------------------------------------------
'�t�@�C����UTF8(BOM����)���𔻒肷��
' path : IN : �t�@�C���p�X(��΃p�X)
' Ret : True/False (True=UTF8(BOM����), False=���L�ȊO)
'-------------------------------------------------------------
Public Function IsUTF8_WithBom(ByVal path As String) As Boolean
    If IsExistsFile(path) = False Then
        Err.Raise 53, , "[IsUTF8_WithBom] �w�肳�ꂽ�t�@�C�������݂��܂��� (path" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsUTF8_WithBom] �p�X���������܂� (path=" & path & ")"
    End If

    Dim bytedata() As Byte: bytedata = ReadBinary(path, 3)
    Dim length As Integer: length = UBound(bytedata) + 1
    
    If length < 3 Then
        IsUTF8_WithBom = False
        Exit Function
    End If
    
    If bytedata(0) = &HEF And bytedata(1) = &HBB And bytedata(2) = &HBF Then
        IsUTF8_WithBom = True
    Else
        IsUTF8_WithBom = False
    End If
    
End Function

'-------------------------------------------------------------
'�t�@�C�����o�C�i���Ƃ��Ďw��T�C�Y�ǂݍ���
' path : IN : �t�@�C���p�X(��΃p�X)
' readsize : IN : �ǂݍ��ރT�C�Y
' Ret : �ǂݍ��񂾃o�C�i���z��
'-------------------------------------------------------------
Public Function ReadBinary(ByVal path As String, ByVal readsize As Integer) As Byte()
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[ReadBinary] �p�X���������܂� (path=" & path & ")"
    End If

    Dim readdata() As Byte
    
    If readsize <= 0 Then
        ReadBinary = readdata()
        Exit Function
    End If
    
    Dim filenum As Integer: filenum = FreeFile
    
    Open path For Binary Access Read As #filenum
    
    ReDim readdata(readsize - 1)
    
    Get #filenum, , readdata
    
    Close #filenum
    
    ReadBinary = readdata
End Function

'-------------------------------------------------------------
'�w��t�H���_�z���Ɏw��g���q�̃t�@�C�������݂��邩
' path : IN : �t�H���_�p�X(��΃p�X)
' in_ext : IN : �g���q(Ex. "*.vb")
' Ret : True/False (True=���݂���, False=���݂��Ȃ�)
'-------------------------------------------------------------
Public Function IsExistsExtensionFile(ByVal path As String, ByVal in_ext As String) As Boolean
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsExistsExtensionFile] �p�X���������܂� (path=" & path & ")"
    End If

    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    Dim file As Object
    Dim Ext As String: Ext = Replace(in_ext, "*", "")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(path)
    
    For Each subfolder In folder.SubFolders
        If IsExistsExtensionFile(subfolder.path, Ext) Then
            Set fso = Nothing
            Set folder = Nothing
            
            IsExistsExtensionFile = True
            Exit Function
        End If
    Next subfolder
    
    For Each file In folder.Files
        If right(file.Name, Len(Ext)) = Ext Then
            Set fso = Nothing
            Set folder = Nothing
        
            IsExistsExtensionFile = True
            Exit Function
        End If
    Next file
    
    Set fso = Nothing
    Set folder = Nothing

    IsExistsExtensionFile = False
End Function

'-------------------------------------------------------------
'���O�t�@�C�����I�[�v������
' logfile_path : IN : ���O�t�@�C���p�X(��΃p�X)
'-------------------------------------------------------------
Public Sub OpenLog(ByVal logfile_path As String)
    If is_log_opened = True Then
        '���łɃI�[�v�����Ă���̂Ŗ���
        Exit Sub
    End If

    If IsMaxOverPath(logfile_path) = True Then
        Err.Raise 53, , "[OpenLog] �p�X���������܂� (logfile_path=" & logfile_path & ")"
    End If

    logfile_num = FreeFile()
    Open logfile_path For Append As logfile_num
    is_log_opened = True
End Sub

'-------------------------------------------------------------
'���O�t�@�C���ɏ�������
' contents : IN : �������ޓ��e
'-------------------------------------------------------------
Public Sub WriteLog(ByVal Contents As String)
    If is_log_opened = False Then
        '�I�[�v������Ă��Ȃ��̂Ŗ���
        Exit Sub
    End If
    Print #logfile_num, Format(Date, "yyyy/mm/dd") & " " & Format(Now, "hh:mm:ss") & ":" & Contents
End Sub

'-------------------------------------------------------------
'���O�t�@�C�����N���[�Y����
'-------------------------------------------------------------
Public Sub CloseLog()
    If is_log_opened = False Then
        '�I�[�v������Ă��Ȃ��̂Ŗ���
        Exit Sub
    End If
    Close logfile_num
    logfile_num = -1
    is_log_opened = False
End Sub

'-------------------------------------------------------------
'�z��̋�s���폜����
' arr : IN : ������z��
' Ret : ��s���폜�����z��
'-------------------------------------------------------------
Public Function DeleteEmptyArray(ByRef arr() As String) As String()
    Dim result() As String
    Dim i As Integer
    Dim count As Integer
    Dim wk As String
    
    If IsEmptyArray(arr) = True Then
        DeleteEmptyArray = result
        Exit Function
    End If
    
    count = 0
    For i = LBound(arr) To UBound(arr)
        wk = Replace(Replace(Replace(arr(i), vbCrLf, ""), vbCr, ""), vbLf, "")
        If wk <> "" Then
            ReDim Preserve result(count)
            result(count) = wk
            count = count + 1
        End If
    Next i
    DeleteEmptyArray = result
End Function

'-------------------------------------------------------------
'�t�@�C�����X�g���쐬����
' path : IN : �t�H���_�p�X(��΃p�X)
' ext : IN : �g���q(Ex."*.vb")
' is_subdir : IN : �T�u�t�H���_�܂ނ� (True=�܂�)
' Ret : �t�@�C�����X�g(��΃p�X�̃��X�g)
'-------------------------------------------------------------
Public Function CreateFileList( _
    ByVal path As String, _
    ByVal Ext As String, _
    ByVal is_subdir As Boolean _
) As String()
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[CreateFileList] �p�X���������܂� (path=" & path & ")"
    End If

    Dim list() As String: list = CreateFileListMain(path, Ext, is_subdir)
    CreateFileList = FilterFileListByExtension(DeleteEmptyArray(list), Ext)
End Function

Private Function CreateFileListMain( _
    ByVal path As String, _
    ByVal Ext As String, _
    ByVal is_subdir As Boolean _
) As String()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim filelist() As String
    Dim cnt As Integer

    Dim file As String, f As Object
    file = Dir(path & "\" & Ext)
    
    If file <> "" Then
        If IsEmptyArray(filelist) = True Then
            cnt = 0
        Else
            cnt = UBound(filelist) + 1
        End If
        
        ReDim Preserve filelist(cnt)
        filelist(cnt) = path & "\" & file
    End If
    
    Do While file <> ""
        file = Dir()
        If file <> "" Then
            cnt = UBound(filelist) + 1
            ReDim Preserve filelist(cnt)
            filelist(cnt) = path & "\" & file
        End If
    Loop
    
    If is_subdir = False Then
        Set fso = Nothing
        CreateFileListMain = filelist
        Exit Function
    End If
    
    Dim filelist_sub() As String
    Dim filelist_merge() As String
    
    For Each f In fso.GetFolder(path).SubFolders
        filelist_sub = CreateFileListMain(f.path, Ext, is_subdir)
        filelist = MergeArray(filelist_sub, filelist)
    Next f
    
    Set fso = Nothing
    CreateFileListMain = filelist
End Function

'-------------------------------------------------------------
'�t�@�C���p�X�̔z�񂩂�w��g���q�̃t�@�C���݂̂�V�����z��ɃR�s�[���ĕԂ��B
' path_list : I : �t�@�C���p�X�̔z��
' in_ext : I : �g���q(Ex. "*.txt")
' Ret : �t�B���^�[��̃t�@�C���p�X�̔z��
'-------------------------------------------------------------
Function FilterFileListByExtension(ByRef path_list() As String, in_ext As String) As String()
    Dim i As Long
    Dim j As Long: j = 0
    Dim filtered_list() As String
    Dim Ext As String: Ext = Replace(in_ext, "*", "")
    
    If in_ext = "*.*" Then
        FilterFileListByExtension = path_list
        Exit Function
    End If
    
    If IsEmptyArray(path_list) = True Then
        FilterFileListByExtension = path_list
        Exit Function
    End If
      
    For i = 0 To UBound(path_list)
        If right(path_list(i), Len(Ext)) = Ext Then
            ReDim Preserve filtered_list(j)
            filtered_list(j) = path_list(i)
            j = j + 1
        End If
    Next i
    
    FilterFileListByExtension = filtered_list
End Function

'-------------------------------------------------------------
'2�̔z����������ĕԂ�
' array1 : IN : �z��1
' array2 : IN : �z��2
' Ret : ���������z��
'-------------------------------------------------------------
Public Function MergeArray(ByRef array1 As Variant, ByRef array2 As Variant) As Variant
    Dim merged As Variant
    merged = Split(Join(array1, vbCrLf) & vbCrLf & Join(array2, vbCrLf), vbCrLf)
    MergeArray = merged
End Function

'-------------------------------------------------------------
'2�̃e�L�X�g�t�@�C�����r���Ĉ�v���Ă��邩��Ԃ�
' file1 : IN : �t�@�C��1�p�X(��΃p�X)
' file2 : IN : �t�@�C��2�p�X(��΃p�X)
' Ret : ��r���� : True/False (True=��v)
'-------------------------------------------------------------
Public Function IsMatchTextFiles(ByVal file1 As String, ByVal file2 As String) As Boolean
    If IsMaxOverPath(file1) = True Or IsMaxOverPath(file2) = True Then
        Err.Raise 53, , "[IsMatchTextFiles] �p�X���������܂� (file1=" & file1 & ", file2=" & file2 & ")"
    End If

    Dim filesize1 As Long: filesize1 = FileLen(file1)
    Dim filesize2 As Long: filesize2 = FileLen(file2)
    
    'TODO:�o�C�i�����x���Ŕ�r���ׂ�
    
    '�܂��t�@�C���T�C�Y�Ń`�F�b�N
    If filesize1 = 0 And filesize2 = 0 Then
        '�ǂ����0byte�Ȃ̂ň�v
        IsMatchTextFiles = True
        Exit Function
    ElseIf filesize1 <> filesize2 Then
        '�t�@�C���T�C�Y���قȂ�̂ŕs��v
        IsMatchTextFiles = False
        Exit Function
    ElseIf filesize1 = 0 Or filesize2 = 0 Then
        '�ǂ��炩��0byte�Ȃ̂ŕs��v
        IsMatchTextFiles = False
        Exit Function
    End If

    Dim fso1, fso2 As Object
    Dim ts1, ts2 As Object
    
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    
    Const READ_ONLY = 1
    Set ts1 = fso1.OpenTextFile(file1, READ_ONLY)
    Set ts2 = fso2.OpenTextFile(file2, READ_ONLY)
    
    Dim contents1 As String: contents1 = ts1.ReadAll
    Dim contents2 As String: contents2 = ts2.ReadAll
    
    ts1.Close
    ts2.Close
    Set ts1 = Nothing
    Set ts2 = Nothing
    Set fso1 = Nothing
    Set fso2 = Nothing
    
    IsMatchTextFiles = (contents1 = contents2)
End Function

'-------------------------------------------------------------
'������̔z��̖����ɕ������ǉ�����
' ary : IN/OUT : ������̔z��
' value : IN : �ǉ����镶����
'-------------------------------------------------------------
Public Sub AppendArray(ByRef ary() As String, ByVal value As String)
    If IsEmptyArray(ary) = True Then
        ReDim Preserve ary(0)
        ary(0) = value
    Else
        Dim cnt As Integer: cnt = UBound(ary) + 1
        ReDim Preserve ary(cnt)
        ary(cnt) = value
    End If
End Sub

Public Sub AppendArrayLong(ByRef ary() As String, ByVal value As String)
    If IsEmptyArrayLong(ary) = True Then
        ReDim Preserve ary(0)
        ary(0) = value
    Else
        Dim cnt As Long: cnt = UBound(ary) + 1
        ReDim Preserve ary(cnt)
        ary(cnt) = value
    End If
End Sub

'-------------------------------------------------------------
'�t�H���_�p�X��񋓂���B�i�T�u�t�H���_�܂ށj
' ���ӁFpath�͖߂�l�ɂ͊܂܂Ȃ�
' path : IN : �t�H���_�p�X�i��΃p�X�j
' Ret : �t�H���_�p�X���X�g
'-------------------------------------------------------------
Public Function GetFolderPathList(ByVal path As String) As String()
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[GetFolderPathList] �p�X���������܂� (path=" & path & ")"
    End If

    Dim fso As Object
    Dim top_dir As Object
    Dim sub_dir As Object
    Dim path_list() As String
    Dim dir_cnt As Long
    Dim i, j As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set top_dir = fso.GetFolder(path)

    dir_cnt = top_dir.SubFolders.count
    If dir_cnt > 0 Then
        ReDim path_list(dir_cnt - 1)
        i = 0
        For Each sub_dir In top_dir.SubFolders
            path_list(i) = sub_dir.path
            i = i + 1
            
            Dim sub_path_list() As String
            sub_path_list = GetFolderPathList(sub_dir.path)
            
            '�T�u�t�H���_���̃p�X��z��ɒǉ�����
            If sub_path_list(0) <> "" Then
                Dim cnt As Integer: cnt = UBound(path_list) + UBound(sub_path_list) + 1
                ReDim Preserve path_list(cnt)
                For j = LBound(sub_path_list) To UBound(sub_path_list)
                    path_list(i) = sub_path_list(j)
                    i = i + 1
                Next j
            End If
        Next sub_dir
        
        GetFolderPathList = path_list
    Else
        Dim ret_empty(0) As String
        GetFolderPathList = ret_empty
    End If
    
    Set sub_dir = Nothing
    Set top_dir = Nothing
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'�t�H���_���R�s�[����(�T�u�t�H���_�܂�)
' src_path : IN : �R�s�[���t�H���_�p�X(��΃p�X)
' dst_path : IN : �R�s�[��t�H���_�p�X(��΃p�X)
'-------------------------------------------------------------
Public Sub CopyFolder(ByVal src_path As String, dest_path As String)
    If IsMaxOverPath(src_path) = True Or IsMaxOverPath(dest_path) = True Then
        Err.Raise 53, , "[CopyFolder] �p�X���������܂� (src_path=" & src_path & ", dest_path=" & dest_path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '�R�s�[���̃t�H���_�����݂��Ȃ��ꍇ�A�G���[�𔭐�������
    If Not fso.FolderExists(src_path) Then
        Err.Raise 53, , "[CopyFolder] �w�肳�ꂽ�t�H���_�����݂��܂���B(src_path=" & src_path & ")"
    End If
    
    '�R�s�[��̃t�H���_�����݂��Ȃ��ꍇ�A�쐬����
    If Not fso.FolderExists(dest_path) Then
        CreateFolder dest_path
    End If
    
    '�R�s�[���̃t�H���_���̃t�@�C�����R�s�[����
    Const OVERWRITE = True
    Dim file As Object
    For Each file In fso.GetFolder(src_path).Files
        fso.CopyFile file.path, fso.BuildPath(dest_path, file.Name), OVERWRITE
    Next
    
    '�R�s�[���̃t�H���_���̃T�u�t�H���_���R�s�[����
    Dim subfolder As Object
    For Each subfolder In fso.GetFolder(src_path).SubFolders
        CopyFolder subfolder.path, fso.BuildPath(dest_path, subfolder.Name)
    Next
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'Yes/No���b�Z�[�W�{�b�N�X��\������
' msg : IN : ���b�Z�[�W
' Ret : True/False (True=Yes, False=No)
'-------------------------------------------------------------
Public Function ShowYesNoMessageBox(ByVal msg As String) As Boolean
    Dim result As Integer: result = MsgBox(msg, vbYesNo, "Confirm")
    
    If result = vbYes Then
        ShowYesNoMessageBox = True
    Else
        ShowYesNoMessageBox = False
    End If
End Function

'-------------------------------------------------------------
'�O���A�v���P�[�V���������s���A�I������܂őҋ@����
' exe_path : IN : �O���A�v���P�[�V����(exe)�̐�΃p�X
'                 exe�ɓn���p�����[�^������ꍇ���ꏏ�ɏ�������
' Ret : �v���Z�X�̖߂�l
'-------------------------------------------------------------
Public Function RunProcessWait(ByVal exe_path As String) As Long
    If IsMaxOverPath(exe_path) = True Then
        Err.Raise 53, , "[RunProcessWait] �p�X���������܂� (exe_path=" & exe_path & ")"
    End If

    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    
    Const NOT_DISP = 0
    Const DISP = 1
    Const WAIT = True
    Const NO_WAIT = False
    
    Dim Process As Object
    Set Process = wsh.Exec(exe_path)
    
    '�v���Z�X�������ɒʒm���󂯎��
    Do While Process.Status = 0
        DoEvents
    Loop
    
    '�v���Z�X�̖߂�l���擾����
    RunProcessWait = Process.ExitCode
    
    Set Process = Nothing
    Set wsh = Nothing
End Function

'-------------------------------------------------------------
' BAT�t�@�C�������s����
' bat_path : IN : BAT�t�@�C���̐�΃p�X
'                 BAT�ɓn���p�����[�^������ꍇ���ꏏ�ɏ�������
' Ret : BAT�̖߂�l(exit /b 0�̏ꍇ0���߂�)
'-------------------------------------------------------------
Public Function RunBatFile(ByVal bat_path As String) As Long
    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    Dim returnValue As Variant
    
    Const NOT_DISP = 0
    Const DISP = 1
    Const WAIT = True
    Const NO_WAIT = False
    
    returnValue = wsh.Run(bat_path, NOT_DISP, WAIT)
    
    RunBatFile = CLng(returnValue)
    
    Set wsh = Nothing
End Function

'-------------------------------------------------------------
'�O��̃_�u���N�H�[�e�[�V�������������ĕԂ�
' ��:"hoge" �� hoge
' target : IN : �Ώە�����
' Ret : ������̕�����
'-------------------------------------------------------------
Public Function RemoveQuotes(ByVal target As String) As String
    '""�ň͂܂�Ă��邩���`�F�b�N
    If left(target, 1) = """" And right(target, 1) = """" Then
        '""���폜���ĕԂ�
        RemoveQuotes = Mid(target, 2, Len(target) - 2)
    Else
        RemoveQuotes = target
    End If
End Function

'-------------------------------------------------------------
'�p�X������̖�����\���������ĕԂ�
' path : IN : �p�X������
' Ret : �p�X������
'-------------------------------------------------------------
Public Function RemoveTrailingBackslash(ByVal path As String) As String
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[RemoveTrailingBackslash] �p�X���������܂� (path=" & path & ")"
    End If

    If right(path, 1) = "\" Then
        path = left(path, Len(path) - 1)
    End If
    RemoveTrailingBackslash = path
End Function

'-------------------------------------------------------------
'�t�@�C���̓��e���w�肳�ꂽ�V�[�g�ɏo�͂���
' file_path : IN : �t�@�C���p�X (��΃p�X)
' sheet_name : IN : �V�[�g��
'-------------------------------------------------------------
Public Sub OutputTextFileToSheet(ByVal file_path As String, ByVal sheet_name As String)
    If IsExistsFile(file_path) = False Or sheet_name = "" Then
        Err.Raise 53, , "[OutputTextFileToSheet] �w�肳�ꂽ�t�@�C�������݂��܂��� (file_path=" & file_path & ")"
    End If

    If IsMaxOverPath(file_path) = True Then
        Err.Raise 53, , "[OutputTextFileToSheet] �p�X���������܂� (file_path=" & file_path & ")"
    End If

    '���[�N�p�ɃR�s�[����
    Dim wk As String: wk = CopyUniqueFile(file_path, "")
    
    '���[�N�t�@�C����SJIS�ɕϊ�����
    UTF8toSJIS wk, False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '�t�@�C�����J��
    Const FORMAT_ASCII = 0
    Const FORMAT_UNICODE = -1
    Const READ_ONLY = 1
    Const IS_CREATE_FILE = False
    Dim fileobj As Object
    Set fileobj = fso.OpenTextFile(wk, READ_ONLY, IS_CREATE_FILE, FORMAT_ASCII)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheet_name)
    
    '�t�@�C���̓��e���V�[�g�ɏo��
    Dim row As Integer: row = 1
    
    Do While Not fileobj.AtEndOfStream
        ws.Cells(row, 1).value = fileobj.ReadLine
        row = row + 1
    Loop
    
    fileobj.Close
    Set fileobj = Nothing
    Set fso = Nothing
    
    '���[�N�t�@�C�����폜����
    DeleteFile wk
End Sub

'-------------------------------------------------------------
'SJIS�Ńe�L�X�g�t�@�C�����쐬����
' contents : IN : ���e
' path : IN : �t�@�C���p�X (��΃p�X)
'-------------------------------------------------------------
Public Sub CreateSJISTextFile(ByRef Contents() As String, ByVal path As String)
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[CreateSJISTextFile] �p�X���������܂� (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim txt As Object
    Dim i As Long
    
    Dim is_overwrite As Boolean: is_overwrite = True
    Dim IS_UNICODE As Boolean: IS_UNICODE = False
    
    Set txt = fso.CreateTextFile(path, is_overwrite, IS_UNICODE)
    
    For i = LBound(Contents) To UBound(Contents)
        txt.WriteLine Contents(i)
    Next i
    
    txt.Close
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'�t�@�C�����쐬����
' path : IN : �t�@�C���p�X (��΃p�X)
'-------------------------------------------------------------
Public Sub CreateFile(ByVal path As String)
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[CreateFile] �p�X���������܂� (path=" & path & ")"
    End If
    
    If path = "" Then
        Err.Raise 53, , "[CreateFile] �p�X���w�肳��Ă��܂��� (path=" & path & ")"
    End If
    
    If IsExistsFile(path) = True Then
        Exit Sub
    End If
    
    Dim dir_path As String
    dir_path = Common.GetFolderPath(path)
    
    If Common.IsExistsFolder(dir_path) = False Then
        CreateFolder dir_path
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.CreateTextFile path
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'�T�u�t�H���_���܂Ƃ߂č쐬����
' path : IN : �t�H���_�p�X (��΃p�X)
'-------------------------------------------------------------
Public Sub CreateFolder(ByVal path As String)
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[CreateFolder] �p�X���������܂� (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folders() As String
    folders = Split(path, Application.PathSeparator)
  
    Dim ary As Variant
    Dim i As Long
    For i = LBound(folders) To UBound(folders)
        ary = folders
        ReDim Preserve ary(i)
        If Not fso.FolderExists(Join(ary, Application.PathSeparator)) Then
            Call fso.CreateFolder(Join(ary, Application.PathSeparator))
        End If
    Next
  
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'�t�H���_���폜����
' path : IN : �t�H���_�p�X (��΃p�X)
'-------------------------------------------------------------
Public Sub DeleteFolder(ByVal path As String)
    If IsExistsFolder(path) = False Then
        Exit Sub
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[DeleteFolder] �p�X���������܂� (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    fso.DeleteFolder path
    
    Set fso = Nothing
End Sub

'-------------------------------------------------------------
'�t�H���_���ړ�����
' src_path : IN : �ړ����t�H���_�p�X (��΃p�X)
' dst_path : IN : �ړ���t�H���_�p�X (��΃p�X)
'-------------------------------------------------------------
Public Sub MoveFolder(ByVal src_path As String, ByVal dst_path As String)
    If IsExistsFolder(src_path) = False Then
        Err.Raise 53, , "[MoveFolder] �ړ����t�H���_�����݂��܂��� (src_path=" & src_path & ")"
        Exit Sub
    End If

    If IsMaxOverPath(src_path) = True Or IsMaxOverPath(dst_path) = True Then
        Err.Raise 53, , "[MoveFolder] �p�X���������܂� (src_path=" & src_path & ", dst_path=" & dst_path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim err_msg As String
    Dim retry As Integer
    For retry = 0 To 3

On Error Resume Next
        fso.MoveFolder src_path, dst_path
    
        err_msg = Err.Description
        Err.Clear
On Error GoTo 0

        If err_msg = "" Then
            Exit For
        End If
        
        WaitSec 1

    Next retry
    
    Set fso = Nothing
    
    If err_msg <> "" Then
        Err.Raise 53, , "[MoveFolder] �G���[! (err_msg=" & err_msg & ")"
    End If
    
End Sub

'-------------------------------------------------------------
'������z��̋��ʕ������Ԃ�
' list : IN : ������z��
' Ret : ���ʕ�����
'       Ex. list = ["hogeAbcdef", "hogeXyz", "hogeApple"]
'           Ret = "hoge"
'-------------------------------------------------------------
Function GetCommonString(ByRef list() As String) As String
    Dim common_string As String
    Dim i, j As Long
    Dim flag As Boolean
    
    '�ŏ��̕���������ʕ�����̏����l�Ƃ���
    common_string = list(0)
    
    '�e��������r����
    For i = 1 To UBound(list)
        flag = False
        '���ʕ������擾����
        For j = 1 To Len(common_string)
            If Mid(common_string, j, 1) <> Mid(list(i), j, 1) Then
                common_string = left(common_string, j - 1)
                flag = True
                Exit For
            End If
        Next j
    Next i
    
    '���ʂ��o�͂���
    GetCommonString = common_string
End Function

'-------------------------------------------------------------
'��΃t�@�C���p�X�̐e�t�H���_�p�X���擾����
' path : IN : �t�@�C���p�X (��΃p�X)
' Ret : �e�t�H���_�p�X (��΃p�X)
'       Ex. path = "C:\tmp\abc.txt"
'           Ret = "C:\tmp"
'-------------------------------------------------------------
Public Function GetFolderNameFromPath(ByVal path As String) As String
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[GetFolderNameFromPath] �p�X���������܂� (path=" & path & ")"
    End If

    Dim last_separator As Long
    
    last_separator = InStrRev(path, Application.PathSeparator)
    
    If last_separator > 0 Then
        GetFolderNameFromPath = left(path, last_separator - 1)
    Else
        GetFolderNameFromPath = path
    End If
End Function

'-------------------------------------------------------------
'���΃p�X���΃p�X�ɕϊ�����
' base_path : IN : ��ƂȂ�t�H���_�p�X(��΃p�X)
' ref_path : IN : �t�@�C���p�X�i���΃p�X)
' Ret : ��΃p�X
'       Ex. base_path = "C:\tmp\abc"
'           ref_path = "..\cdf\xyz.txt"
'           Ret = "C:\tmp\cdf\xyz.txt"
'-------------------------------------------------------------
Public Function GetAbsolutePathName(ByVal base_path As String, ByVal ref_path As String) As String
    If IsMaxOverPath(base_path) = True Or IsMaxOverPath(ref_path) = True Then
        Err.Raise 53, , "[GetAbsolutePathName] �p�X���������܂� (base_path=" & base_path & ", ref_path=" & ref_path & ")"
    End If

     Dim fso As Object
     Set fso = CreateObject("Scripting.FileSystemObject")
     
     GetAbsolutePathName = fso.GetAbsolutePathName(fso.BuildPath(base_path, ref_path))
     
     Set fso = Nothing
End Function

'-------------------------------------------------------------
'�t�@�C���̑��݃`�F�b�N���s��
' path : IN : �t�@�C���p�X(��΃p�X)
' Ret : True/False (True=���݂���)
'-------------------------------------------------------------
Public Function IsExistsFile(ByVal path As String) As Boolean
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsExistsFile] �p�X���������܂� (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(path) Then
        IsExistsFile = True
    Else
        IsExistsFile = False
    End If
    
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'�t�H���_�̑��݃`�F�b�N���s��
' path : IN : �t�H���_�p�X(��΃p�X)
' Ret : True/False (True=���݂���)
'-------------------------------------------------------------
Public Function IsExistsFolder(ByVal path As String) As Boolean
    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[IsExistsFolder] �p�X���������܂� (path=" & path & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(path) Then
        IsExistsFolder = True
    Else
        IsExistsFolder = False
    End If
    
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'�t�@�C��������g���q��Ԃ�
' filename : IN : �t�@�C����
' isRaw    : IN : True=�擾�����g���q�����̂܂ܕԂ��AFalse=�������ɕϊ����ĕԂ�(�f�t�H���g)
' Ret : �t�@�C�����̊g���q
'        Ex. "abc.txt"�̏ꍇ�A"txt"���Ԃ�
'            "."���܂܂�Ă��Ȃ��ꍇ��""���Ԃ�
'-------------------------------------------------------------
Public Function GetFileExtension(ByVal filename As String, Optional ByVal isRaw As Boolean = False) As String
    Dim dot_pos As Integer
    
    ' "."�̈ʒu���擾
    dot_pos = InStrRev(filename, ".")
    
    ' �g���q���擾
    If dot_pos > 0 Then
        If isRaw = False Then
            GetFileExtension = LCase(right(filename, Len(filename) - dot_pos))
        Else
            GetFileExtension = right(filename, Len(filename) - dot_pos)
        End If
    Else
        GetFileExtension = ""
    End If
End Function

'-------------------------------------------------------------
'�w��t�H���_�z�����w��t�@�C�����Ō������Ă��̓��e��Ԃ�
' target_folder : IN :�����t�H���_�p�X(��΃p�X)
' target_file : IN :�����t�@�C����
' Ret : �ǂݍ��񂾃t�@�C���̓��e
'       �z��̖����ɂ͌����t�@�C���̐�΃p�X���i�[����
'-------------------------------------------------------------
Public Function SearchAndReadFiles(ByVal target_folder As String, ByVal target_file As String) As String()
    If IsMaxOverPath(target_folder) = True Then
        Err.Raise 53, , "[SearchAndReadFiles] �p�X���������܂� (target_folder=" & target_folder & ")"
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(target_folder)
    
    Dim fileobj As Object
    For Each fileobj In folder.Files
        If fso.FileExists(fileobj.path) And fso.GetFileName(fileobj.path) Like target_file Then
            '�����Ώۂ̃t�@�C����ǂݍ���
            Dim Contents As String: Contents = ReadTextFileBySJIS(fileobj.path)

            '�t�@�C���̓��e��z��Ɋi�[����
            Dim lines() As String: lines = Split(Contents, vbCrLf)
            
            '�����Ƀt�@�C���p�X��ǉ�����
            Dim lines_cnt As Integer: lines_cnt = UBound(lines)
            ReDim Preserve lines(lines_cnt + 1)
            lines(lines_cnt + 1) = file.path
            SearchAndReadFiles = lines
            Set fileobj = Nothing
            Set fso = Nothing
            Exit Function
        End If
    Next file
    
    '�T�u�t�H���_����������
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        Dim result() As String
        result = SearchAndReadFiles(subfolder.path, target_file, is_sjis)
        If UBound(result) >= 1 Then
            '�T�u�t�H���_���猋�ʂ��Ԃ��Ă����ꍇ�́A���̌��ʂ�Ԃ�
            SearchAndReadFiles = result
            Set fso = Nothing
            Exit Function
        End If
    Next subfolder
    
    '�����Ώۂ̃t�@�C����������Ȃ������ꍇ�́A��̔z���Ԃ�
    Dim ret_empty(0) As String
    SearchAndReadFiles = ret_empty
    Set fso = Nothing
End Function

'-------------------------------------------------------------
'SJIS�Ńe�L�X�g�t�@�C����ǂݍ���
'��UTF8�̃t�@�C����SJIS�ɕϊ����ēǂݍ���!
' path : IN : �t�@�C���p�X (��΃p�X)
' Ret : �ǂݍ��񂾓��e
'-------------------------------------------------------------
Public Function ReadTextFileBySJIS(ByVal path As String) As String
    If IsExistsFile(path) = False Then
        Err.Raise 53, , "[ReadTextFileBySJIS] �w�肳�ꂽ�t�@�C�������݂��܂��� (path=" & path & ")"
    End If

    If IsMaxOverPath(path) = True Then
        Err.Raise 53, , "[ReadTextFileBySJIS] �p�X���������܂� (path=" & path & ")"
    End If

    '���[�N�p�ɃR�s�[����
    Dim wk As String: wk = CopyUniqueFile(path, "")
    
    '���[�N�t�@�C����SJIS�ɕϊ�����
    UTF8toSJIS wk, False
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Const FORMAT_ASCII = 0
    Const READ_ONLY = 1
    Const IS_CREATE_FILE = False
    
    Dim fileobj As Object
    Set fileobj = fso.OpenTextFile(wk, READ_ONLY, IS_CREATE_FILE, FORMAT_ASCII)
    Dim Contents As String: Contents = fileobj.ReadAll
    
    fileobj.Close
    Set fileobj = Nothing
    Set fso = Nothing
    
    '���[�N�t�@�C�����폜����
    DeleteFile wk
    
    ReadTextFileBySJIS = RTrim(Contents)
End Function

'-------------------------------------------------------------
'UTF-8�`���̃e�L�X�g�t�@�C����ǂݍ���
' file_path : IN : �t�@�C���p�X (��΃p�X)
' Ret : �ǂݍ��񂾓��e
'-------------------------------------------------------------
Public Function ReadTextFileByUTF8(ByVal file_path) As String
    If IsMaxOverPath(file_path) = True Then
        Err.Raise 53, , "[ReadTextFileByUTF8] �p�X���������܂� (file_path=" & file_path & ")"
    End If
    
    Dim Contents As String
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile file_path
        Contents = .ReadText
        .Close
    End With
    
    ReadTextFileByUTF8 = Contents
End Function

'-------------------------------------------------------------
'�z�񂪋󂩂��`�F�b�N����
' arr : IN : �z��
' Ret : True/False (True=��)
'-------------------------------------------------------------
Public Function IsEmptyArray(arr As Variant) As Boolean
    On Error Resume Next
    Dim i As Integer
    i = UBound(arr)
    If i >= 0 And Err.Number = 0 Then
        IsEmptyArray = False
    Else
        IsEmptyArray = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function IsEmptyArrayLong(arr As Variant) As Boolean
    On Error Resume Next
    Dim i As Long
    i = UBound(arr)
    If i >= 0 And Err.Number = 0 Then
        IsEmptyArrayLong = False
    Else
        IsEmptyArrayLong = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

'-------------------------------------------------------------
'n�b�҂�
' sec : I : �҂���(�b) ����������
'-------------------------------------------------------------
Public Sub WaitSec(ByVal sec As Double)
    Application.WAIT [Now()] + sec / 86400
End Sub

'-------------------------------------------------------------
'���ݓ������~���b�P�ʂ̕�����ŕԂ�
' Ret :Ex."20230326123456001"
'-------------------------------------------------------------
Public Function GetNowTimeString() As String
    Dim t As SYSTEMTIME

    Call GetLocalTime(t)
    
    Dim yyyy As String: yyyy = Format(t.wYear, "0000")
    Dim mm As String: mm = Format(t.wMonth, "00")
    Dim dd As String: dd = Format(t.wDay, "00")
    Dim hh As String: hh = Format(t.wHour, "00")
    Dim mn As String: mn = Format(t.wMinute, "00")
    Dim ss As String: ss = Format(t.wSecond, "00")
    Dim fff As String: fff = Format(t.wMilliseconds, "000")
    
    GetNowTimeString = yyyy & mm & dd & hh & mn & ss & fff
End Function

Public Function GetNowTimeString_OLD() As String
    Dim str_date As String
    Dim str_time As String
    
    str_date = Format(Date, "yyyymmdd")
    str_time = Format(Time, "hhmmss")
    
    GetNowTimeString_OLD = str_date & str_time
End Function

'-------------------------------------------------------------
'�V�[�g�̑��݃`�F�b�N
' wb : I : ���[�N�u�b�N
' sheet_name : I : �V�[�g��
' Ret : True/False (True=���݂���)
'-------------------------------------------------------------
Public Function IsExistSheet(ByRef wb As Workbook, ByVal sheet_name As String) As Boolean
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        If ws.Name = sheet_name Then
            IsExistSheet = True
            Exit Function
        End If
    Next ws
    
    IsExistSheet = False
End Function

'-------------------------------------------------------------
'�V�[�g���폜����
' wb : I : ���[�N�u�b�N
' sheet_name : I : �V�[�g��
'-------------------------------------------------------------
Public Sub DeleteSheet(ByRef wb As Workbook, ByVal sheet_name As String)
    If IsExistSheet(wb, sheet_name) = False Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    wb.Sheets(sheet_name).Delete
    Application.DisplayAlerts = True
End Sub

'-------------------------------------------------------------
'�V�[�g��ǉ�����
' wb : I : ���[�N�u�b�N
' sheet_name : I : �V�[�g��
'-------------------------------------------------------------
Public Sub AddSheet(ByRef wb As Workbook, ByVal sheet_name As String)
    DeleteSheet wb, sheet_name
    wb.Worksheets.Add.Name = sheet_name
End Sub

'-------------------------------------------------------------
'�u�b�N���A�N�e�B�u�ɂ���
' book_name : IN : �u�b�N��(Excel�t�@�C����)
'-------------------------------------------------------------
Public Sub ActiveBook(ByVal book_name As String)
    If IsOpenWorkbook(book_name) = False Then
        Err.Raise 53, , "[ActiveBook] �u�b�N���J����Ă��܂��� (book_name=" & book_name & ")"
    End If
    
    Dim wb As Workbook
    Set wb = Workbooks(book_name)
    wb.Activate
End Sub

'-------------------------------------------------------------
'�w�肳�ꂽ�V�[�g�̎w��Z���ɒl���o�͂���
' book_name : IN : ���[�N�u�b�N
' sheet_name : IN : �V�[�g��
' cell_row : �s
' cell_clm : ��
' contents : IN : �o�͂�����e
'-------------------------------------------------------------
Public Sub UpdateSheet( _
    ByRef book_name As Workbook, _
    ByVal sheet_name As String, _
    ByVal cell_row As Long, ByVal cell_clm As Long, _
    ByVal Contents As String)
    
    If IsExistSheet(book_name, sheet_name) = False Then
        Err.Raise 53, , "[UpdateSheet] �V�[�g��������܂��� (book_name=" & book_name & "), sheet_name=" & sheet_name & ")"
    End If
    
    If cell_row < 0 Or cell_clm < 0 Then
        Err.Raise 53, , "[UpdateSheet] �Z���ʒu���s���ł� (cell_row=" & cell_row & "), cell_clm=" & cell_clm & ")"
    End If
    
    Dim ws As Worksheet
    Set ws = book_name.Sheets(sheet_name)
    
    ws.Cells(cell_row, cell_clm).value = Contents
End Sub

