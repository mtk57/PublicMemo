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
    Const START_ROW = 16
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
        row = row + 1
    Loop

    Common.WriteLog "CheckAndCollectParam E"
End Sub

'Sub Param�����Ɏ��s���Ă���
Private Sub ExecSubParam()
    Common.WriteLog "ExecSubParam S"
    
    If Common.IsEmptyArray(sub_params) = True Then
        Err.Raise 53, , "�L����Sub param������܂���"
    End If

    Dim i As Integer
    Dim copy_datas() As CopyData
    Dim sub_param As SubParam
    
    For i = LBound(sub_params) To UBound(sub_params)
        Set sub_param = sub_params(i)
        
        Common.WriteLog "��Main Loop (SubParam Row#=" & sub_param.GetSubParamRowNumber() & ")"
        
        '�]�L���f�[�^�����W����
        copy_datas = CollectSrcDatas(sub_param)
            
        If Common.IsEmptyArray(copy_datas) = True Then
            Common.WriteLog "�]�L���f�[�^������܂���B"
            GoTo CONTINUE_FOR
        End If
            
        '�]�L���f�[�^��]�L��ɓ]�L����
        Transcription sub_param, copy_datas

CONTINUE_FOR:

    Next i
    
    Common.WriteLog "ExecSubParam E"
End Sub

'�]�L���f�[�^�����W����
Private Function CollectSrcDatas(ByRef sub_param As SubParam) As CopyData()
    Common.WriteLog "CollectSrcDatas S"

    Dim ws As Worksheet
    Dim copy_datas() As CopyData
    Dim copy_data As CopyData
    Dim cnt As Long
    Dim cell As Range

    'SRC�t�@�C���p�X��SRC�V�[�g�����J��
    Const READ_ONLY_FLG = True
    Const VISIBLE_FLG = True
    Set ws = Common.GetSheet( _
                sub_param.GetSrcFilePath(), _
                sub_param.GetSrcSheetName(), _
                READ_ONLY_FLG, _
                VISIBLE_FLG _
             )
    
    'SRC������̉��F�Z���ɑΉ�����SRC�]�L��̒l�����W����
    Dim key_rng As Range
    Dim value_rng As Range
    Dim key_clm As String: key_clm = sub_param.GetSrcFindClm()
    Dim val_clm As String: val_clm = sub_param.GetSrcTranClm()
    
    Set key_rng = ws.Range(key_clm & "1:" & key_clm & ws.Cells(ws.Rows.count, key_clm).End(xlUp).row)
    Set value_rng = ws.Range(val_clm & "1:" & val_clm & ws.Cells(ws.Rows.count, val_clm).End(xlUp).row)

    cnt = 0
    For Each cell In key_rng
        '���W�Ώۂ͉��F�Z���݂̂Ƃ���
        If cell.Interior.Color = RGB(255, 255, 0) Then
            ReDim Preserve copy_datas(cnt)
            Set copy_data = New CopyData
            copy_data.Init cell.value, value_rng.Cells(cell.row, 1).value
            Set copy_datas(cnt) = copy_data
            cnt = cnt + 1
        End If
    Next cell
    
    If main_param.IsNotClose() = False Then
        'SRC�t�@�C�������
        Common.CloseBook (Common.GetFileName(sub_param.GetSrcFilePath()))
    End If
    
    CollectSrcDatas = copy_datas
    
    Common.WriteLog "CollectSrcDatas E"
End Function

'�]�L����
Private Sub Transcription(ByRef sub_param As SubParam, ByRef copy_datas() As CopyData)
    Common.WriteLog "Transcription S"
    
    Dim ws As Worksheet
    Dim book_name As String
    Dim row As Long
    Dim keyword As String
    Dim found_row As Long
    Dim trans_rng As Range
    Dim copy_data As CopyData
    
    'DST�t�@�C���p�X��DST�V�[�g�����J��
    Const READ_ONLY_FLG = False
    Const VISIBLE_FLG = True
    Set ws = Common.GetSheet( _
                sub_param.GetDstFilePath(), _
                sub_param.GetDstSheetName(), _
                READ_ONLY_FLG, _
                VISIBLE_FLG _
             )
    book_name = Common.GetFileName(sub_param.GetDstFilePath())
    
    Dim last_row As Long: last_row = Common.GetLastRowFromWorksheet(ws, sub_param.GetDstFindClm())
    
    'SRC������̒l���ADST������ɂ��邩��������
    '����΁ASRC�]�L��̒l��DST�]�L��ɓ����
    For row = LBound(copy_datas, 1) To UBound(copy_datas, 1)
    
        Set copy_data = copy_datas(row)
        keyword = copy_data.GetKey()
        
        If keyword = "" Then
            GoTo CONTINUE_ROW
        End If
        
        Dim find_row As Long: find_row = 1
        
        Do
            '�w���̑S�s���w�胏�[�h�Ō������A�q�b�g�����s�ԍ����擾����
            found_row = Common.FindRowByKeywordFromWorksheet( _
                           ws, _
                           sub_param.GetDstFindClm(), _
                           find_row, _
                           keyword _
                        )
        
            If found_row = 0 Then
                '������Ȃ�!
                'Common.WriteLog "Search keyword is not found!" & vbCrLf & _
                '                "row=" & row & vbCrLf & _
                '                "keyword=" & keyword
                '����
                Exit Do
            End If
            
            '���������̂œ]�L
            Set trans_rng = ws.Range(sub_param.GetDstTranClm() & found_row)
            trans_rng.value = copy_data.GetValue()
            
            If last_row = found_row Then
                '�ŏI�s�Ȃ̂Ń��[�v�𔲂���
                Exit Do
            End If
            
            '���������s�̎��s���Č���
            find_row = found_row + 1
        
        Loop
        
CONTINUE_ROW:
        
    Next row
    
    If main_param.IsNotClose() = False Then
        'DST�t�@�C����ۑ����ĕ���
        Common.SaveAndCloseBook (book_name)
    End If
    
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

Function GetYellowCellData( _
  ByVal filePath As String, _
  ByVal sheetName As String, _
  ByVal searchCol As String, _
  ByVal dataCol As String _
  ) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim searchRange As Range
    Dim dataRange As Range
    Dim cell As Range
    Dim result() As String
    Dim i As Long

    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Sheets(sheetName)

    Set searchRange = ws.Range(searchCol & "1:" & searchCol & ws.Cells(ws.Rows.count, searchCol).End(xlUp).row)
    Set dataRange = ws.Range(dataCol & "1:" & dataCol & ws.Cells(ws.Rows.count, dataCol).End(xlUp).row)

    ReDim result(0 To searchRange.Cells.count - 1, 0 To 1)

    i = 0
    For Each cell In searchRange
        If cell.Interior.Color = RGB(255, 255, 0) Then
            result(i, 0) = cell.value
            result(i, 1) = dataRange.Cells(cell.row, 1).value
            i = i + 1
        End If
    Next cell

    wb.Close SaveChanges:=False

    GetYellowCellData = result
End Function
