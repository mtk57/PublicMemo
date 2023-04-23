Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String

'�p�����[�^
Private main_param As MainParam
Private sub_params() As SubParam

'���C������
Public Sub Run()
    Common.WriteLog "Run S"

    Worksheets("main").Activate
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    CheckAndCollectParam
    
    'Sub Param�����Ɏ��s���Ă���
    ExecSubParam
    
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
    
    Common.WriteLog main_param.GetAllValue()

    'Sub Params
    Const START_ROW = 13
    Const SUB_ROWS = 1 + (5 * 2)    '1=ENABLE, 5=SubParam, 2=SRC, DST
    Dim row As Long: row = START_ROW
    Dim cnt As Long: cnt = 0
    
    Do
        Dim sub_param As SubParam
        Set sub_param = New SubParam
        
        Common.WriteLog "row=" & row
        sub_param.Init row
        sub_param.Validate

        Common.WriteLog sub_param.GetAllValue()
        
        If sub_param.GetEnable() = "STOPPER" Then
            Exit Do
        ElseIf sub_param.GetEnable() = "DISABLE" Then
            GoTo CONTINUE
        End If
        
        ReDim Preserve sub_params(cnt)
        Set sub_params(cnt) = sub_param
        cnt = cnt + 1
        
CONTINUE:
        row = row + SUB_ROWS + 1
    Loop

    Common.WriteLog "CheckAndCollectParam E"
End Sub

'Sub Param�����Ɏ��s���Ă���
Private Sub ExecSubParam()
    Common.WriteLog "ExecSubParam S"
    
    If Common.IsEmptyArray(sub_params) = True Or _
       UBound(sub_params) < 0 Then
        Err.Raise 53, , "�L����Sub param������܂���"
    End If

    Dim i As Integer
    Dim src_datas() As String
    
    For i = LBound(sub_params) To UBound(sub_params)
        Common.WriteLog "��Main Loop (i=" & i & ")"
    
        Dim sub_param As SubParam
        Set sub_param = sub_params(i)
        
        '�]�L���f�[�^�����W����
        src_datas = CollectSrcDatas(sub_param)
            
        '���W�f�[�^��]�L����
        Transcription sub_param, src_datas

    Next i
    
    Common.WriteLog "ExecSubParam E"
End Sub

'�]�L���f�[�^�����W����
Private Function CollectSrcDatas(ByRef sub_param As SubParam) As String()
    Common.WriteLog "CollectSrcDatas S"

    Dim ws As Worksheet
    Dim book_name As String
    Dim temp_sheet As String

    'SRC�t�@�C���p�X��SRC�V�[�g�����J��
    Set ws = Common.GetSheet(sub_param.GetSrcFilePath(), sub_param.GetSrcSheetName(), False)
    book_name = Common.GetFileName(sub_param.GetSrcFilePath())
    
    '��Ɨp�V�[�g��ǉ�����
    Common.ActiveBook book_name
    temp_sheet = Common.GetNowTimeString()
    ActiveWorkbook.Worksheets.Add.name = temp_sheet
    
    'SRC�J�n�s����ASRC������ASRC�]�L��̑S�s����Ɨp�V�[�g�ɃR�s�[����
    'TODO:SRC������̕����s�Ή�
    CopyColumnToAnotherSheet _
      sub_param.GetSrcSheetName(), sub_param.GetSrcFindClm(), sub_param.GetSrcStartRow(), _
      temp_sheet, "A", 1
    CopyColumnToAnotherSheet _
      sub_param.GetSrcSheetName(), sub_param.GetSrcTranClm(), sub_param.GetSrcStartRow(), _
      temp_sheet, "B", 1
    
    'SRC�J�n�s����ASRC������ASRC�]�L��̑S�s��2�����z��Ɋi�[����
    CollectSrcDatas = Common.GetSheetContentsByStringArray(temp_sheet)
    
    'SRC�t�@�C���p�X�ƕ���
    Common.CloseBook (Common.GetFileName(sub_param.GetSrcFilePath()))
    
    Common.WriteLog "CollectSrcDatas E"
End Function

'�]�L����
Private Sub Transcription(ByRef sub_param As SubParam, ByRef src_datas() As String)
    Common.WriteLog "Transcription S"
    
    Dim ws As Worksheet
    Dim book_name As String
    Dim r As Long, c As Long
    Dim find_word As String
    Dim found_row As Long
    
    'DST�t�@�C���p�X��DST�V�[�g�����J��
    Set ws = Common.GetSheet(sub_param.GetDstFilePath(), sub_param.GetDstSheetName(), True)
    book_name = Common.GetFileName(sub_param.GetDstFilePath())
    
    'SRC������̒l���ADST������ɂ��邩��������
    '����΁ASRC�]�L��̒l��DST�]�L��ɓ����
    For r = LBound(src_datas, 1) To UBound(src_datas, 1)
    
        find_word = Trim(src_datas(r, 1))
        
        If find_word = "" Then
            GoTo CONTINUE_ROW
        End If
        
        '�w���̑S�s���w�胏�[�h�Ō������A�q�b�g�����s�ԍ����擾����
        found_row = Common.FindRowByKeyword( _
                       ws, _
                       sub_param.GetDstFindClm(), _
                       sub_param.GetDstStartRow(), _
                       find_word _
                    )
    
        If found_row = 0 Then
            '������Ȃ�!
            Common.WriteLog "Search value is not found!" & vbCrLf & _
                            "r=" & r & vbCrLf & _
                            "find_word=" & find_word
            'TODO:�������񖳎�
            GoTo CONTINUE_ROW
        End If
        
        
        
CONTINUE_ROW:
        
    Next r
    
    'DST�t�@�C���p�X�ƕ���
    Common.CloseBook (Common.GetFileName(sub_param.GetDstFilePath()))
    
    Common.WriteLog "Transcription E"
End Sub

'�R�s�[���͈̔͂��R�s�[��͈̔͂ɃR�s�[
Private Sub CopyColumnToAnotherSheet( _
  ByVal src_sheet_name As String, _
  ByVal src_clm As String, _
  ByVal src_start_row As Long, _
  ByVal dst_sheet_name As String, _
  ByVal dst_clm As String, _
  ByVal dst_start_row As Long _
  )
    Common.WriteLog "CopyColumnToAnotherSheet S"
    
    Dim src_sheet As Worksheet
    Dim dst_sheet As Worksheet
    Dim last_row As Long
    Dim src_range As Range
    Dim dst_range As Range
    
    Set src_sheet = ActiveWorkbook.Worksheets(src_sheet_name)
    Set dst_sheet = ActiveWorkbook.Worksheets(dst_sheet_name)
    
    last_row = src_sheet.Cells(Rows.count, src_clm).End(xlUp).row
    
    Set src_range = src_sheet.Range(src_clm & src_start_row & ":" & src_clm & last_row)
    Set dst_range = dst_sheet.Range(dst_clm & dst_start_row & ":" & dst_clm & last_row)
    
    '�R�s�[���͈̔͂��R�s�[��͈̔͂ɃR�s�[
    src_range.Copy dst_range
    
    Common.WriteLog "CopyColumnToAnotherSheet E"
End Sub


