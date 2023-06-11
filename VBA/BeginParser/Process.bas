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

Private parse_datas() As String

'--------------------------------------------------------
'���C������
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    Erase targets
    Erase results
    Erase parse_datas
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
            
            '�t�@�C���p�X�Ŏn�܂�A" Begin "���܂ލs�������W����
            If Common.IsMatchByRegExp(line, "^[a-zA-Z]:\\", True) = True And _
               InStr(line, " Begin ") > 0 Then
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
    Dim begin_codes() As String
    
    Set result = New ParseResult
    result.Init target, main_param.GetFormatType()
    
    '����������VB�R�[�h�̃p�[�X����
    
    'Begin�`End�܂ł̃R�[�h���擾����
    GetBeginCodes result

    If result.GetBeginCodesCount() = 0 Then
        'Begin�͌��o������End�����o�o���Ȃ������ꍇ(���ʒu���Y���Ă���\����)
        Common.WriteLog "ParseCode E1"
        Exit Sub
    End If

    '1�s���p�[�X���āA���ʃI�u�W�F�N�g���쐬����
    ParseBeginCode result
    
    parse_datas = Common.DeleteEmptyArray(parse_datas)
    result.SetBeginMembers parse_datas
    Erase parse_datas
    
    If result.GetBeginMembersCount() = 0 Then
        'Begin�`End�͌��o���������\�b�h�E�v���p�e�B�����g�p�̏ꍇ
        Common.WriteLog "ParseCode E2"
        Exit Sub
    End If
    
    ReDim Preserve results(result_cnt)
    Set results(result_cnt) = result
    result_cnt = result_cnt + 1
    
    Common.WriteLog "ParseCode E"
End Sub

'Begin�`End�܂ł̃R�[�h��z��ŕԂ�
'������q��Begin�͖�������
Private Sub GetBeginCodes( _
    ByRef result As ParseResult _
)
    Common.WriteLog "GetBeginCodes S"
    
    Dim raw_contents() As String
    Dim begin_codes() As String
    Dim i As Long
    Dim cnt As Long: cnt = 0
    Dim line As String
    Dim ext As String: ext = result.GetExtension()
    Dim is_find As Boolean: is_find = False
    Dim clm_wk As Long
    Dim first_clm As Long: first_clm = 0
    Dim is_ignore As Boolean: is_ignore = False
    
    '�t�@�C���p�X�̃t�@�C�����J��
    raw_contents = GetTargetContents(result)
    
    'Begin�`End�܂ł̍s��z��ɓ����
    For i = result.GetRowNum() - 1 To UBound(raw_contents)
        line = raw_contents(i)

        If Common.IsCommentCode(line, ext) = True Then
            '�R�����g�s�Ȃ̂Ŏ��̍s��
            GoTo CONTINUE
        End If

        '�E�R�����g���������Ă���
        line = Common.RemoveRightComment(line, ext)
        
        If Common.IsMatchByRegExp(line, "^Begin .*$", True) = True Then
        
            'Begin�����o
            
            clm_wk = Common.FindFirstCasePosition(line)
            
            If first_clm = 0 Then
                '�ŏ��Ɍ��o����Begin�̌��ʒu��ێ����Ă���
                first_clm = clm_wk
            End If
        
        ElseIf Common.IsMatchByRegExp(line, "^End$", True) = True Then
        
            'End Begin�����o
            
            clm_wk = Common.FindFirstCasePosition(line)
            If clm_wk = first_clm Then
                'Grep���ʂ�Begin�ɑΉ�����End�𔭌������̂ŏI��
                ReDim Preserve begin_codes(cnt)
                begin_codes(cnt) = line
                is_find = True
                Exit For
            End If
                    
        End If
    
        ReDim Preserve begin_codes(cnt)
        begin_codes(cnt) = line
        cnt = cnt + 1

CONTINUE:
    
    Next i
    
    If is_find = False Then
        Dim err_msg As String
    
        err_msg = "Grep���ʂ�Begin�ɑΉ�����End��������܂��� (target=" & result.GetTarget() & ")"
    
        If Common.ShowYesNoMessageBox( _
            "[GetBeginCodes]�ŃG���[���������܂����B�����𑱍s���܂���?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[GetBeginCodes] �G���[! (err_msg=" & err_msg & ")"
        End If
        
        Common.WriteLog err_msg
        Common.WriteLog "GetBeginCodes E1"
        Exit Sub
    End If
    
    result.SetBeginCodes begin_codes
    
    Common.WriteLog "GetBeginCodes E"
End Sub

'1�s���p�[�X���āA���ʃI�u�W�F�N�g���쐬����
Private Sub ParseBeginCode( _
    ByRef result As ParseResult _
)
    Common.WriteLog "ParseBeginCode S"

    Parse result, result.GetBeginCodes()(0), result.GetBeginClass()

    Common.WriteLog "ParseBeginCode E"
End Sub

'�K�w�T��
Private Sub Parse( _
    ByRef result As ParseResult, _
    ByVal name As String, _
    ByVal path As String _
)
    Common.WriteLog "Parse S"

    AppendPropertyPath result, name, path
    
    Dim sub_begins() As String
    sub_begins = GetSubBeginList(result, name)
    
    If Common.IsEmptyArray(sub_begins) = True Then
        Common.WriteLog "Parse E1"
        Exit Sub
    End If
    
    Dim i As Long
    Dim sub_name As String
    Dim sub_path As String
    
    For i = 0 To UBound(sub_begins)
        sub_name = sub_begins(i)
        sub_path = path & "/" & Replace(Replace(Trim(sub_name), "BeginProperty", ""), "Begin", "")
        
        '�ċA
        Parse result, sub_name, sub_path
        
    Next i

    Common.WriteLog "Parse E"
End Sub

'���݂̊K�w�ɂ���v���p�e�B(Key = Value)��񋓂��āA�����p�X�̖����Ƀf���~�^�t���Ō�������
Private Sub AppendPropertyPath( _
    ByRef result As ParseResult, _
    ByVal name As String, _
    ByVal path As String _
)
    Common.WriteLog "AppendPropertyPath S"

    Const REG = "^(?!.*(Begin |BeginProperty |EndProperty$|End$)).*$"
    Dim contents() As String
    contents = GetSubBeginContents(result, name)
    
    Dim i As Long
    Dim line As String
    Dim member() As String
    Dim cnt As Long: cnt = 0
    
    '�v���p�e�B�������W
    For i = 0 To UBound(contents)
        line = contents(i)
        
        If Common.IsMatchByRegExp(line, REG, True) = True Then
            ReDim Preserve member(cnt)
            member(cnt) = line
            cnt = cnt + 1
        End If
    Next i

    '�ŏ����ʒu���擾
    Dim min_clm As Long: min_clm = GetMinColumn(member)
    
    '���݂̊K�w�̃v���p�e�B�����ɂ���
    Dim member_current() As String
    Dim clm_wk As Long
    cnt = 0
    
    For i = 0 To UBound(member)
        line = member(i)
        
        clm_wk = Common.FindFirstCasePosition(line)
        
        If clm_wk <= min_clm Then
            ReDim Preserve member_current(cnt)
            member_current(cnt) = path & "/" & Trim(line)
            cnt = cnt + 1
        End If
    Next i
    
    '�Ō�Ƀ}�[�W
    parse_datas = Common.MergeArray(parse_datas, member_current)

    Common.WriteLog "AppendPropertyPath E"
End Sub
    
'�ŏ����ʒu��Ԃ�
'TODO: Common��
Private Function GetMinColumn(ByRef ary() As String) As Long
    Common.WriteLog "GetMinColumn S"

    Dim i As Long
    Dim clm_wk As Long
    Dim line As String
    Dim min_clm As Long: min_clm = -1
    
    For i = 0 To UBound(ary)
        line = ary(i)
        clm_wk = Common.FindFirstCasePosition(line)
        
        If min_clm = -1 Then
            min_clm = clm_wk
        Else
            If min_clm > clm_wk Then
                '�ŏ��𔭌�
                min_clm = clm_wk
            End If
        End If
        
        
    Next i
    
    GetMinColumn = min_clm

    Common.WriteLog "GetMinColumn E"
End Function

'���݂̊K�w�ɂ���T�u�K�w("Begin" or "BeginPrpperty"�Ŏn�܂�)��񋓂���
Private Function GetSubBeginList( _
    ByRef result As ParseResult, _
    ByVal name As String _
) As String()
    Common.WriteLog "GetSubBeginList S"

    Const REG = "Begin |BeginProperty "
    
    Dim contents() As String
    contents = GetSubBeginContents(result, name)
    
    Dim i As Long
    Dim line As String
    Dim member() As String
    Dim cnt As Long: cnt = 0
    
    Common.WriteLog "contents=" & CStr(UBound(contents))
    
    'Begin, BeginProperty���������W
    For i = 0 To UBound(contents)
        If i = 0 Then
            '1�s�ڂ͖���
            GoTo CONTINUE
        End If
        
        line = contents(i)
        
        If Common.IsMatchByRegExp(line, REG, True) = True Then
            ReDim Preserve member(cnt)
            member(cnt) = line
            cnt = cnt + 1
        End If
CONTINUE:
    Next i
    
    If Common.IsEmptyArray(member) = True Then
        GetSubBeginList = member
        Common.WriteLog "GetSubBeginList E1"
        Exit Function
    End If
    
    '�ŏ��������擾
    Dim min_clm As Long: min_clm = GetMinColumn(member)
    
    '���݂̊K�w��Begin, BeginProperty�����ɂ���
    Dim member_current() As String
    Dim clm_wk As Long
    cnt = 0
    
    For i = 0 To UBound(member)
        line = member(i)
        
        clm_wk = Common.FindFirstCasePosition(line)
        
        If clm_wk <= min_clm Then
            ReDim Preserve member_current(cnt)
            member_current(cnt) = line
            cnt = cnt + 1
        End If
    Next i
    
    GetSubBeginList = member_current

    Common.WriteLog "GetSubBeginList E"
End Function

'�w�肳�ꂽ�K�w�̓��e��Ԃ�
Private Function GetSubBeginContents( _
    ByRef result As ParseResult, _
    ByVal name As String _
) As String()
    Common.WriteLog "GetSubBeginContents S"

    Dim i As Long
    Dim line As String
    Dim ext As String: ext = result.GetExtension()
    Dim first_clm As Long: first_clm = -1
    Dim clm_wk As Long
    Dim is_find As Boolean: is_find = False
    Dim contents() As String
    Dim cnt As Long: cnt = 0
    Dim end_word As String
    
    Dim begin_type As Integer   '0=Begin, 1=BeginProperty
    
    If Left(Trim(name), 6) = "Begin " Then
        begin_type = 0
        end_word = "^End$"
    ElseIf Left(Trim(name), 14) = "BeginProperty " Then
        begin_type = 1
        end_word = "^EndProperty$"
    Else
        Err.Raise 53, , "�L�[���[�h��������܂���! (target=" & result.GetTarget() & ")"
    End If
    
    'Begin�`End�܂ł̍s��z��ɓ����
    For i = 0 To UBound(result.GetBeginCodes())
        line = result.GetBeginCodes()(i)
        
        If Common.IsCommentCode(line, ext) = True Then
            '�R�����g�s�Ȃ̂Ŏ��̍s��
            GoTo CONTINUE
        End If
    
        '�E�R�����g���������Ă���
        line = Common.RemoveRightComment(line, ext)
        
        If first_clm = -1 And line = name Then
            first_clm = Common.FindFirstCasePosition(line)
        End If
    
        If first_clm = -1 Then
            '�Ώۂ������Ă��Ȃ��̂Ŗ���
            GoTo CONTINUE
        End If
    
        If Common.IsMatchByRegExp(Trim(line), end_word, True) = True Then
            clm_wk = Common.FindFirstCasePosition(line)
            If clm_wk = first_clm Then
                'Begin�ɑ΂���End�𔭌������̂ŏI��
                ReDim Preserve contents(cnt)
                contents(cnt) = line
                is_find = True
                Exit For
            End If
        End If
    
        ReDim Preserve contents(cnt)
        contents(cnt) = line
        cnt = cnt + 1
    
CONTINUE:
    
    Next i
    
    If is_find = False Then
        Dim err_msg As String
        
        err_msg = "Grep���ʂ�Begin�ɑΉ�����End��������܂��� (target=" & result.GetTarget() & ")"
        
        If Common.ShowYesNoMessageBox( _
            "[GetSubBeginContents]�ŃG���[���������܂����B�����𑱍s���܂���?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[GetSubBeginContents] �G���[! (err_msg=" & err_msg & ")"
        End If
        
        Common.WriteLog err_msg
        Common.WriteLog "GetSubBeginContents E1"
        Exit Function
    End If
    
    GetSubBeginContents = contents

    Common.WriteLog "GetSubBeginContents E"
End Function

'�V�[�g�Ɍ��ʂ��o�͂���
Private Sub OutputSheet()
    Common.WriteLog "OutputSheet S"
    
    If Common.IsEmptyArray(results) = True Then
        Common.WriteLog "OutputSheet E1"
        Exit Sub
    End If
    
    '�V�[�g��ǉ�
    Dim sheet_name As String: sheet_name = Common.GetNowTimeString()
    Common.AddSheet ActiveWorkbook, sheet_name
    
    '�V�[�g�̃^�C�g����ǉ�
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("A1").value = "Begin��̉�͌���"
    
    '�񖼂�ǉ�
    ws.Range("A3").value = "GREP����"
    ws.Range("B3").value = "�t�@�C���p�X"
    ws.Range("C3").value = "�v���p�e�B"
    ws.Range("D3").value = "�l"
    ws.Range("E3").value = "���[�g"
    ws.Range("F3").value = "�K�w1"
    ws.Range("G3").value = "�K�w2"
    ws.Range("H3").value = "�K�w3"
    ws.Range("I3").value = "�K�w4"
    ws.Range("J3").value = "�K�w5"
    ws.Range("K3").value = "�K�w6"
    ws.Range("L3").value = "�K�w7"
    ws.Range("M3").value = "�K�w8"
    ws.Range("N3").value = "�K�w9"
    ws.Range("O3").value = "�K�w10"
    
    
    Const START_ROW = 4
    Dim i As Long
    Dim j As Long
    Dim offset_row As Long: offset_row = 0
    Dim result As ParseResult
    Dim row As Long: row = START_ROW
    Dim members() As String
    
    Dim cnt As Long
    
    Dim key_ As String
    Dim val_ As String
    Dim k1 As String
    Dim k2 As String
    Dim k3 As String
    Dim k4 As String
    Dim k5 As String
    Dim k6 As String
    Dim k7 As String
    Dim k8 As String
    Dim k9 As String
    
    Dim key_val() As String
    
    
    '���ʃI�u�W�F�N�g���X�g�Ń��[�v
    For i = 0 To UBound(results)
        Set result = results(i)
        
        If result.GetBeginMembersCount() = 0 Then
            GoTo CONTINUE
        End If
        
        '���ʃI�u�W�F�N�g�̓��e���L��
        members = result.GetBeginMembers()
        
        For j = 0 To UBound(members)
            Dim items() As String: items = Split(members(j), "/")
            
            If Common.IsEmptyArray(items) = True Then
                GoTo CONTINUE_J
            End If
            
            cnt = UBound(items)
            
            key_ = ""
            val_ = ""
            k1 = ""
            k2 = ""
            k3 = ""
            k4 = ""
            k5 = ""
            k6 = ""
            k7 = ""
            k8 = ""
            k9 = ""
        
            ws.Cells(row + i + offset_row + j, 1).value = result.GetTarget()
            ws.Cells(row + i + offset_row + j, 2).value = result.GetFilePath()
            
            key_val = Split(items(cnt), "=")
            
            If Common.IsEmptyArray(key_val) = True Or UBound(key_val) = 0 Then
                GoTo CONTINUE_J
            End If
            
            key_ = Trim(key_val(0))
            val_ = Trim(key_val(1))
            
            If cnt = 1 Then
            
            ElseIf cnt = 2 Then
                k1 = items(1)
            ElseIf cnt = 3 Then
                k1 = items(1)
                k2 = items(2)
            ElseIf cnt = 4 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
            ElseIf cnt = 5 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
                k4 = items(4)
            ElseIf cnt = 6 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
                k4 = items(4)
                k5 = items(5)
            ElseIf cnt = 7 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
                k4 = items(4)
                k5 = items(5)
                k6 = items(6)
            ElseIf cnt = 8 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
                k4 = items(4)
                k5 = items(5)
                k6 = items(6)
                k7 = items(7)
            ElseIf cnt = 9 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
                k4 = items(4)
                k5 = items(5)
                k6 = items(6)
                k7 = items(7)
                k8 = items(8)
            ElseIf cnt = 10 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
                k4 = items(4)
                k5 = items(5)
                k6 = items(6)
                k7 = items(7)
                k8 = items(8)
                k9 = items(9)
            End If
            
            ws.Cells(row + i + offset_row + j, 3).value = key_
            ws.Cells(row + i + offset_row + j, 4).value = val_
            ws.Cells(row + i + offset_row + j, 5).value = result.GetBeginClass()
            
            ws.Cells(row + i + offset_row + j, 6).value = k1
            ws.Cells(row + i + offset_row + j, 7).value = k2
            ws.Cells(row + i + offset_row + j, 8).value = k3
            ws.Cells(row + i + offset_row + j, 9).value = k4
            ws.Cells(row + i + offset_row + j, 10).value = k5
            ws.Cells(row + i + offset_row + j, 11).value = k6
            ws.Cells(row + i + offset_row + j, 12).value = k7
            ws.Cells(row + i + offset_row + j, 13).value = k8
            ws.Cells(row + i + offset_row + j, 14).value = k9
        
CONTINUE_J:
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
    Dim contents() As String
    
    '�t�@�C�����J���āA�S�s��z��Ɋi�[����
    If result.GetEncode() = "SJIS" Then
        raw_contents = Common.ReadTextFileBySJIS(result.GetFilePath())
    ElseIf result.GetEncode() = "UTF-8" Then
        raw_contents = Common.ReadTextFileByUTF8(result.GetFilePath())
    Else
        Dim err_msg As String: err_msg = "���T�|�[�g�̃G���R�[�h�ł�" & vbCrLf & _
                  "path=" & result.GetFilePath()
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

