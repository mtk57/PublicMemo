Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String

Private main_param As MainParam
Private conditions() As String
Private contents_cache() As String
Private before_path As String

Private Const NOT_FOUND = "Not Found."

'--------------------------------------------------------
'���C������
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    CheckAndCollectParam

    '�����������擾����
    Call CreateFindConditions
    
    If Common.IsEmptyArray(conditions) = True Then
        Err.Raise 53, , "������������ł�"
    End If
    

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(main_param.GetTargetSheetName())

    Dim i As Long: i = 0
    Dim j As Long: j = 0
    Dim c As String: c = main_param.GetGrepResultCellColumn()
    Dim c1 As String: c1 = main_param.GetFindWordCellColumn()
    Dim r As Long: r = main_param.GetGrepResultCellRow()
    
    Dim grep_result As String
    Dim target_path As String
    Dim find_word As String
    Dim condition As String
    
    before_path = ""
    Erase contents_cache

    '���C�����[�v
    Do
        grep_result = ws.Range(c & r + i).value
        
        If grep_result = "" Then
            'Grep���ʂ���Ȃ̂ŏI��
            Exit Do
        End If
        
        'Grep���ʂ���t�@�C���p�X���擾����
        target_path = GetFilePathFromGrepResult(grep_result)
        
        '�������[�h���擾����
        find_word = ws.Range(c1 & r + i).value
        
        If target_path <> before_path Then
            before_path = ""
            Erase contents_cache
        End If
        
        If target_path = "" Or Common.IsExistsFile(target_path) = False Or find_word = "" Then
            Call UpdateNotFound(i)
            GoTo SKIP
        End If
        
        '���������������[�v
        For j = 0 To UBound(conditions)
            Dim c2 As Integer: c2 = Common.GetColNumFromA1(main_param.GetFindConditionCell())
        
            '���������Ɍ������[�h��}������
            condition = Replace(conditions(j), main_param.GetReplaceChar(), find_word)
            
            '�t�@�C�����J���āA�������[�h�����������ʂ��擾����
            Dim find_result As String
            find_result = FindRowStringByConditionFromFile(target_path, condition)
            
            '���ʂ��Z���ɏ�������
            ws.Cells(r + i, c2 + j).value = find_result
        Next j
   
SKIP:

        i = i + 1
    Loop

    Set ws = Nothing

    Common.WriteLog "Run E"
End Sub

'�p�����[�^�̃`�F�b�N�Ǝ��W���s��
Private Sub CheckAndCollectParam()
    Common.WriteLog "CheckAndCollectParam S"
    
    Dim err_msg As String

    Set main_param = New MainParam
    main_param.Init

    'Main Params
    main_param.Validate
    
    Common.WriteLog main_param.GetAllValue()
    
    Common.WriteLog "CheckAndCollectParam E"
End Sub

Private Sub CreateFindConditions()
    Common.WriteLog "CreateFindConditions S"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(main_param.GetTargetSheetName())
    
    Dim condition As String
    Dim c As Integer: c = Common.GetColNumFromA1(main_param.GetFindConditionCell())
    Dim r As Long: r = main_param.GetFindConditionCellRow()
    Dim i As Long
    i = 0
    
    Do
        condition = ws.Cells(c + i, r).value
        If condition = "" Then
            '������������Ȃ̂ŏI��
            Exit Do
        End If
        
        ReDim Preserve conditions(i)
        conditions(i) = condition
        
        i = i + 1
    Loop
    
    Set ws = Nothing
    
    Common.WriteLog "CreateFindConditions E"
End Sub

Private Function GetFilePathFromGrepResult(ByVal grep_result As String) As String
    Common.WriteLog "GetFilePathFromGrepResult S (" & grep_result & ")"
    
    Dim tmp() As String
    Dim tmp_path As String
    
    tmp = Common.GetMatchByRegExp(grep_result, "^[A-Z]:.*:", False)
    
    If Common.IsEmptyArray(tmp) = True Then
        'sakura��Grep���ʂł͖����Ǝv����̂ŋ��Ԃ�
        Common.WriteLog "GetFilePathFromGrepResult E-1"
        GetFilePathFromGrepResult = ""
        Exit Function
    End If
    
    tmp_path = tmp(0)
    
    GetFilePathFromGrepResult = Common.ReplaceByRegExp(tmp_path, "\(\d,\d\) *\[.*\]:", "", False)
    
    Common.WriteLog "GetFilePathFromGrepResult E"
End Function

Private Sub UpdateNotFound(ByVal now_row As Long)
    Common.WriteLog "UpdateNotFound S (" & now_row & ")"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(main_param.GetTargetSheetName())
    
    Dim c As Integer: c = Common.GetColNumFromA1(main_param.GetFindConditionCell())
    ws.Cells(c, now_row).value = NOT_FOUND
    
    Set ws = Nothing
    
    Common.WriteLog "UpdateNotFound E"
End Sub

Private Function FindRowStringByConditionFromFile(ByVal path As String, ByVal condition As String) As String
    Common.WriteLog "FindRowStringByConditionFromFile S"
    
    If before_path = "" Then
        '1���
        before_path = path
        contents_cache = GetTargetContents(path)
    End If
    
    If Common.IsEmptyArray(contents_cache) = True Then
        FindRowStringByConditionFromFile = NOT_FOUND
        Common.WriteLog "FindRowStringByConditionFromFile E-1"
        Exit Function
    End If
    
    Dim find_row_num  As Long: find_row_num = Common.FindRowByKeywordFromArray(condition, contents_cache, main_param.IsRegEx())
    
    If find_row_num < 0 Then
        FindRowStringByConditionFromFile = NOT_FOUND
        Common.WriteLog "FindRowStringByConditionFromFile E-2"
        Exit Function
    End If
    
    FindRowStringByConditionFromFile = contents_cache(find_row_num)
    
    Common.WriteLog "FindRowStringByConditionFromFile E"
End Function

'�Ώۃt�@�C����ǂݍ���œ��e��z��ŕԂ�
Private Function GetTargetContents(ByVal path As String) As String()
    Common.WriteLog "GetTargetContents S"
    
    Dim raw_contents As String
    Dim contents() As String
    
    '�t�@�C�����J���āA�S�s��z��Ɋi�[����
    If Common.IsSJIS(path) = True Then
        raw_contents = Common.ReadTextFileBySJIS(path)
    ElseIf Common.IsUTF8(path) = True Then
        raw_contents = Common.ReadTextFileByUTF8(path)
    Else
        Dim err_msg As String: err_msg = "���T�|�[�g�̃G���R�[�h�ł�" & vbCrLf & _
                  "path=" & path
        Common.WriteLog "[GetTargetContents] �����G���[! err_msg=" & err_msg
        
        If Common.ShowYesNoMessageBox( _
            "[GetTargetContents]�ŃG���[���������܂����B�����𑱍s���܂���?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[GetTargetContents] �G���[! (err_msg=" & err_msg & ")"
        End If
        
        Common.WriteLog "GetTargetContents E1"
        GetTargetContents = contents
        Exit Function
    End If
    
    contents = Split(raw_contents, vbCrLf)
    
    GetTargetContents = contents

    Common.WriteLog "GetTargetContents E"
End Function
