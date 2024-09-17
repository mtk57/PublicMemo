Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String

'�p�����[�^
Private main_param As MainParam
Private sub_param As SubParam

Private targets() As String
Private results() As ParseResult
Private result_cnt As Long

'--------------------------------------------------------
'���C������
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    Erase targets
    Erase results
    result_cnt = 0

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    CheckAndCollectParam
    
    '���C�����[�v
    Dim i As Long
    For i = LBound(targets) To UBound(targets)
        Dim target As String: target = targets(i)
        Common.WriteLog "i=" & i & ":[" & target & "]"
    
        '�R�[�h����͂���
        ParseCode target
    Next i
    
    '�V�[�g�Ɍ��ʂ��o�͂���
    OutputSheet

    Common.WriteLog "Run E"
End Sub

'�p�����[�^�̃`�F�b�N�Ǝ��W���s��
Private Sub CheckAndCollectParam()
    Common.WriteLog "CheckAndCollectParam S"
    
    Dim err_msg As String
    
    'Main Params
    Set main_param = New MainParam
    main_param.Init
    main_param.Validate

    'Sub Params
    Set sub_param = New SubParam
    sub_param.Init
    sub_param.Validate

    Common.WriteLog main_param.GetAllValue()
    
    Dim grep_result() As String
    Dim i As Long: i = 0
    Dim cnt As Long: cnt = 0
    Dim line As String
    
    grep_result = sub_param.GetGrepResults()
    
    If main_param.GetFormatType() = "sakura" Then
        For i = LBound(grep_result) To UBound(grep_result)
            line = grep_result(i)
            
            '�t�@�C���p�X�Ŏn�܂�A" With "���܂ލs�������W����
            If Common.IsMatchByRegExp(line, "^[a-zA-Z]:\\", True) = True And _
               InStr(line, " With ") > 0 Then
                ReDim Preserve targets(cnt)
                targets(cnt) = line
                
                cnt = cnt + 1
            End If
        Next i
    End If
    
    Common.WriteLog "CheckAndCollectParam E"
End Sub

'�R�[�h����͂���
Private Sub ParseCode( _
    ByVal target As String _
)
    Common.WriteLog "ParseCode S"
    
    Dim result As ParseResult
    Dim with_codes() As String
    
    Set result = New ParseResult
    result.Init target, main_param.GetFormatType()
    
    '����������VB�R�[�h�̃p�[�X����
    
    'With�`End With�܂ł̃R�[�h���擾����
    GetWithCodes result

    If result.GetWithCodesCount() = 0 Then
        'With�͌��o������End With�����o�o���Ȃ������ꍇ(���ʒu���Y���Ă���\����)
        Common.WriteLog "ParseCode E1"
        Exit Sub
    End If

    '1�s���p�[�X���āA���ʃI�u�W�F�N�g���쐬����
    ParseWithCode result
    
    If result.GetWithMembersCount() = 0 Then
        'With�`End With�͌��o���������\�b�h�E�v���p�e�B�����g�p�̏ꍇ
        Common.WriteLog "ParseCode E2"
        Exit Sub
    End If
    
    ReDim Preserve results(result_cnt)
    Set results(result_cnt) = result
    result_cnt = result_cnt + 1
    
    Common.WriteLog "ParseCode E"
End Sub

'With�`End With�܂ł̃R�[�h��z��ŕԂ�
'������q��With�͖�������
Private Sub GetWithCodes( _
    ByRef result As ParseResult _
)
    Common.WriteLog "GetWithCodes S"
    
    Dim raw_contents() As String
    Dim with_codes() As String
    Dim i As Long
    Dim cnt As Long: cnt = 0
    Dim line As String
    Dim Ext As String: Ext = result.GetExtension()
    Dim is_find As Boolean: is_find = False
    Dim clm_wk As Long
    Dim first_clm As Long: first_clm = 0
    Dim is_ignore As Boolean: is_ignore = False
    
    '�t�@�C���p�X�̃t�@�C�����J��
    raw_contents = GetTargetContents(result)
    
    'With�`End With�܂ł̍s��z��ɓ����
    For i = result.GetRowNum() - 1 To UBound(raw_contents)
        line = raw_contents(i)

        If Common.IsCommentCode(line, Ext) = True Then
            '�R�����g�s�Ȃ̂Ŏ��̍s��
            GoTo CONTINUE
        End If

        '�E�R�����g���������Ă���
        line = Common.RemoveRightComment(line, Ext)
        
        If Common.IsMatchByRegExp(line, "^ *With .*$", True) = True Then
        
            'With�����o
            
            clm_wk = Common.FindFirstCasePosition(line)
            
            If first_clm = 0 Then
                '�ŏ��Ɍ��o����With�̌��ʒu��ێ����Ă���
                first_clm = clm_wk
            End If
            
            If clm_wk <> first_clm Then
                '����q��With�����o�����̂Ŗ���
                is_ignore = True
                GoTo CONTINUE
            End If
        
        ElseIf Common.IsMatchByRegExp(line, "^ *End With$", True) = True Then
        
            'End With�����o
            
            If is_ignore = True Then
                '����q��With�̏I�������o
                is_ignore = False
                GoTo CONTINUE
            End If
        
            clm_wk = Common.FindFirstCasePosition(line)
            If clm_wk = first_clm Then
                'Grep���ʂ�With�ɑΉ�����End With�𔭌������̂ŏI��
                ReDim Preserve with_codes(cnt)
                with_codes(cnt) = line
                is_find = True
                Exit For
            Else
                '�Y���̌덷�����e
                Dim diff As Integer: diff = Max(Abs(first_clm - clm_wk), Abs(clm_wk - first_clm))
                
                If main_param.GetCheckLen() >= diff Then
                    'Grep���ʂ�With�ɑΉ�����End With�𔭌������̂ŏI��
                    ReDim Preserve with_codes(cnt)
                    with_codes(cnt) = line
                    is_find = True
                    Exit For
                End If
            End If
        
        Else
            
            'With, End With�ȊO�̍s
            If is_ignore = True Then
                '����q��With�̏I�������o���Ă��Ȃ��̂Ŗ���
                GoTo CONTINUE
            End If
                    
        End If
    
        ReDim Preserve with_codes(cnt)
        with_codes(cnt) = line
        cnt = cnt + 1

CONTINUE:
    
    Next i
    
    If is_find = False Then
        Dim err_msg As String
    
        err_msg = "��Grep���ʂ�With�ɑΉ�����End With��������܂��� (target=" & result.GetTarget() & ")"
    
        Common.WriteLog err_msg
    
        If main_param.IsConfirm() = True Then
            If Common.ShowYesNoMessageBox( _
                "[GetWithCodes]�ŃG���[���������܂����B�����𑱍s���܂���?" & vbCrLf & _
                "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[GetWithCodes] �G���[! (err_msg=" & err_msg & ")"
            End If
        End If
        
        Common.WriteLog "GetWithCodes E1"
        Exit Sub
    End If
    
    result.SetWithCodes with_codes
    
    Common.WriteLog "GetWithCodes E"
End Sub

'1�s���p�[�X���āA���ʃI�u�W�F�N�g���쐬����
Private Sub ParseWithCode( _
    ByRef result As ParseResult _
)
    Common.WriteLog "ParseWithCode S"
    
    Const MEMBER = "(\s|\()\.[a-zA-Z][a-zA-Z0-9_]*"
    
    Dim i As Long
    Dim j As Long
    Dim with_codes() As String
    Dim with_class As String
    Dim temp_ary() As String
    Dim with_members() As String
    Dim line As String
    
    with_codes = result.GetWithCodes()
    
    For i = 0 To UBound(with_codes)
        line = with_codes(i)
        
        'Common.WriteLog "[" & i & "]=" & line
        
        If i = 0 Then
            with_class = Trim(Replace(line, "With", ""))
            GoTo CONTINUE
        End If
        
        temp_ary = Common.DeleteEmptyArray(Common.GetMatchByRegExp(line, MEMBER, True))
        If Common.IsEmptyArrayLong(temp_ary) = True Then
            '�h�b�g�Ŏn�܂郁�\�b�h�E�v���p�e�B�����݂��Ȃ��̂Ŏ��̍s��
            GoTo CONTINUE
        End If
        
        For j = 0 To UBound(temp_ary)
            temp_ary(j) = Replace(Trim(temp_ary(j)), "(", "")
        Next j
        
        with_members = Common.MergeArray(with_members, temp_ary)
    
CONTINUE:
    
    Next i
    
    If Common.IsEmptyArrayLong(with_members) = True Then
        Common.WriteLog "ParseWithCode E1"
        Exit Sub
    End If
    
    result.SetWithClass with_class
    result.SetWithMembers Common.SortAndDistinctArray(Common.DeleteEmptyArray(with_members))

    Common.WriteLog "ParseWithCode E"
End Sub

'�V�[�g�Ɍ��ʂ��o�͂���
Private Sub OutputSheet()
    Common.WriteLog "OutputSheet S"
    
    If Common.IsEmptyArrayLong(results) = True Then
        Common.WriteLog "OutputSheet E1"
        Exit Sub
    End If
    
    '�V�[�g��ǉ�
    Dim sheet_name As String: sheet_name = Common.GetNowTimeString()
    Common.AddSheet ActiveWorkbook, sheet_name
    
    '�V�[�g�̃^�C�g����ǉ�
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("A1").value = "With��̉�͌���"
    
    '�񖼂�ǉ�
    ws.Range("A3").value = "GREP����"
    ws.Range("B3").value = "�t�@�C���p�X"
    ws.Range("C3").value = "�N���X"
    ws.Range("D3").value = "���\�b�h/�v���p�e�B"
    
    
    Const START_ROW = 4
    Dim i As Long
    Dim j As Long
    Dim offset_row As Long: offset_row = 0
    Dim result As ParseResult
    Dim row As Long: row = START_ROW
    Dim members() As String
    
    '���ʃI�u�W�F�N�g���X�g�Ń��[�v
    For i = 0 To UBound(results)
        Set result = results(i)
        
        If result.GetWithMembersCount() = 0 Then
            GoTo CONTINUE
        End If
        
        '���ʃI�u�W�F�N�g�̓��e���L��
        members = result.GetWithMembers()
        
        For j = 0 To UBound(members)
            
            If j = 0 Then
                ws.Cells(row + i + offset_row + j, 1).Font.Color = RGB(0, 0, 0)
                ws.Cells(row + i + offset_row + j, 2).Font.Color = RGB(0, 0, 0)
                ws.Cells(row + i + offset_row + j, 3).Font.Color = RGB(0, 0, 0)
                ws.Cells(row + i + offset_row + j, 4).Font.Color = RGB(0, 0, 0)
            Else
                ws.Cells(row + i + offset_row + j, 1).Font.Color = RGB(192, 192, 192)
                ws.Cells(row + i + offset_row + j, 2).Font.Color = RGB(192, 192, 192)
                ws.Cells(row + i + offset_row + j, 3).Font.Color = RGB(192, 192, 192)
                ws.Cells(row + i + offset_row + j, 4).Font.Color = RGB(0, 0, 0)
            End If
        
            ws.Cells(row + i + offset_row + j, 1).value = result.GetTarget()
            ws.Cells(row + i + offset_row + j, 2).value = result.GetFilePath()
            ws.Cells(row + i + offset_row + j, 3).value = result.GetWithClass()
            ws.Cells(row + i + offset_row + j, 4).value = members(j)

        Next j
        
        offset_row = offset_row + UBound(members)
        
CONTINUE:
    Next i
    
    Common.WriteLog "OutputSheet E"
End Sub

'�Ώۃt�@�C����ǂݍ���œ��e��z��ŕԂ�
Private Function GetTargetContents( _
    ByRef result As ParseResult _
) As String()
    Common.WriteLog "GetTargetContents S"
    
    Dim raw_contents As String
    Dim Contents() As String
    
    '�t�@�C�����J���āA�S�s��z��Ɋi�[����
    If result.GetEncode() = "SJIS" Then
        raw_contents = Common.ReadTextFileBySJIS(result.GetFilePath())
    ElseIf result.GetEncode() = "UTF-8" Then
        raw_contents = Common.ReadTextFileByUTF8(result.GetFilePath())
    Else
        Dim err_msg As String: err_msg = "���T�|�[�g�̃G���R�[�h�ł�" & vbCrLf & _
                  "path=" & result.GetFilePath()
        Common.WriteLog "[GetTargetContents] �����G���[! err_msg=" & err_msg
        
        If main_param.IsConfirm() = True Then
            If Common.ShowYesNoMessageBox( _
                "[GetTargetContents]�ŃG���[���������܂����B�����𑱍s���܂���?" & vbCrLf & _
                "err_msg=" & err_msg _
            ) = False Then
                Err.Raise 53, , "[GetTargetContents] �G���[! (err_msg=" & err_msg & ")"
            End If
        End If
        
        Common.WriteLog "GetTargetContents E1"
        GetTargetContents = Contents
        Exit Function
    End If
    
    Contents = Split(raw_contents, vbCrLf)
    
    GetTargetContents = Contents

    Common.WriteLog "GetTargetContents E"
End Function

Private Function MySplit(ByVal Expression As String, Optional ByVal Delimiter As String = " ") As Variant
    Dim result() As String
    Dim startPos As Long
    Dim delimPos As Long
    Dim itemCount As Long
    
    ' �z���������
    ReDim result(0 To 0)
    itemCount = 0
    startPos = 1
    
    ' ��������J��Ԃ�����
    Do
        ' ��؂蕶���̈ʒu��������
        delimPos = InStr(startPos, Expression, Delimiter)
        
        If delimPos = 0 Then
            ' �Ō�̍��ڂ�ǉ�
            ReDim Preserve result(0 To itemCount)
            result(itemCount) = Mid(Expression, startPos)
            Exit Do
        Else
            ' ���ڂ�ǉ�
            ReDim Preserve result(0 To itemCount)
            result(itemCount) = Mid(Expression, startPos, delimPos - startPos)
            itemCount = itemCount + 1
            startPos = delimPos + Len(Delimiter)
        End If
    Loop
    
    MySplit = result
End Function

Sub AnalyzeFileContent()
    Dim fso As Object
    Dim textFile As Object
    Dim content As String
    Dim lines() As String
    Dim i As Long
    Dim filePath As String
    
    ' �t�@�C���p�X��ݒ�i�K�X�ύX���Ă��������j
    filePath = "C:\path\to\your\file.txt"
    
    ' FileSystemObject���쐬
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' �t�@�C�����J��
    Set textFile = fso.OpenTextFile(filePath, 1) ' 1 = �ǂݎ���p
    
    ' �t�@�C���̓��e��ǂݍ���
    content = textFile.ReadAll
    
    ' �t�@�C�������
    textFile.Close
    
    ' ���e��CRLF�ŕ���
    lines = Split(content, vbCrLf)
    
    ' ���͌��ʂ�\��
    Debug.Print "���o�C�g��: " & Len(content)
    Debug.Print "Split��̍s��: " & UBound(lines) - LBound(lines) + 1
    
    ' �e�s�̏ڍׂ𕪐�
    For i = LBound(lines) To UBound(lines)
        If Len(lines(i)) = 0 Then
            Debug.Print "�s " & i + 1 & ": ��s"
        Else
            Debug.Print "�s " & i + 1 & ": ���� " & Len(lines(i)) & " ����, ASCII: " & AsciiDump(lines(i))
        End If
    Next i
    
    ' �I�u�W�F�N�g�̉��
    Set textFile = Nothing
    Set fso = Nothing
End Sub

Function AsciiDump(str As String) As String
    Dim i As Long
    Dim result As String
    For i = 1 To Len(str)
        result = result & " " & Right("0" & Hex(Asc(Mid(str, i, 1))), 2)
    Next i
    AsciiDump = Trim(result)
End Function

Option Explicit

Function ReadFileToArray(ByVal filePath As String) As String()
    Dim fso As Object
    Dim textFile As Object
    Dim lines() As String
    Dim lineCount As Long
    Dim tempLine As String
    
    ' FileSystemObject���쐬
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' �t�@�C�������݂��邩�m�F
    If Not fso.FileExists(filePath) Then
        MsgBox "�w�肳�ꂽ�t�@�C����������܂���: " & filePath, vbExclamation
        Exit Function
    End If
    
    ' �t�@�C�����J��
    Set textFile = fso.OpenTextFile(filePath, 1) ' 1 = �ǂݎ���p
    
    ' �z���������
    ReDim lines(0 To 0)
    lineCount = 0
    
    ' �t�@�C����1�s���ǂݍ���
    Do Until textFile.AtEndOfStream
        tempLine = textFile.ReadLine
        ReDim Preserve lines(0 To lineCount)
        lines(lineCount) = tempLine
        lineCount = lineCount + 1
    Loop
    
    ' �t�@�C�������
    textFile.Close
    
    ' �I�u�W�F�N�g�̉��
    Set textFile = Nothing
    Set fso = Nothing
    
    ' ���ʂ�Ԃ�
    ReadFileToArray = lines
End Function

Sub TestReadFileToArray()
    Dim filePath As String
    Dim fileLines() As String
    Dim i As Long
    
    ' �e�X�g�p�̃t�@�C���p�X�i�K�X�ύX���Ă��������j
    filePath = "C:\path\to\your\file.txt"
    
    ' �֐����Ăяo��
    fileLines = ReadFileToArray(filePath)
    
    ' ���ʂ�\��
    Debug.Print "�t�@�C�����̍s��: " & UBound(fileLines) + 1
    
    For i = 0 To UBound(fileLines)
        Debug.Print "�s " & i + 1 & ": " & fileLines(i)
    Next i
End Sub
