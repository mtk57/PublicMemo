Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String

'�p�����[�^
Private main_param As MainParam
Private sub_param As SubParam

Private targets() As String
Public results() As GrepResultInfoStruct
    
'--------------------------------------------------------
'���C������
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)


    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    CheckAndCollectParam
    
    
    '���C��
    results = Common.GetMethodInfoFromGrepResult( _
                targets, _
                main_param.GetFormatType(), _
                main_param.GetLang() _
              )
    
    
    '�V�[�g�Ɍ��ʂ��o�͂���
    Call OutputSheet

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
    
    If main_param.GetFormatType() = GrepAppEnum.sakura Then
        For i = LBound(grep_result) To UBound(grep_result)
            line = grep_result(i)
            
            '�t�@�C���p�X�Ŏn�܂�s�������W����
            If Common.IsMatchByRegExp(line, "^[a-zA-Z]:\\", True) = True Then
                ReDim Preserve targets(cnt)
                targets(cnt) = line
                
                cnt = cnt + 1
            End If
        Next i
    End If
    
    Common.WriteLog "CheckAndCollectParam E"
End Sub

'�V�[�g�Ɍ��ʂ��o�͂���
Private Sub OutputSheet()
    Common.WriteLog "OutputSheet S"
        
    '�V�[�g��ǉ�
    Dim sheet_name As String: sheet_name = Common.GetNowTimeString()
    Common.AddSheet ActiveWorkbook, sheet_name
    
    Dim ws As Worksheet
    Set ws = ActiveSheet

    '�񖼂�ǉ�
    ws.Range("A3").value = "GREP����(Raw)"
    ws.Range("B3").value = "�t�@�C���p�X"
    ws.Range("C3").value = "GREP����"
    ws.Range("D3").value = "�G���[���"
    
    ws.Range("F3").value = "�V�O�l�`��(Raw)"
    ws.Range("G3").value = "���\�b�h��"
    ws.Range("H3").value = "�߂�l"
    ws.Range("I3").value = "����1"
    ws.Range("J3").value = "����2"
    ws.Range("K3").value = "����3"
    ws.Range("L3").value = "����4"
    ws.Range("M3").value = "����5"
    ws.Range("N3").value = "����6"
    ws.Range("O3").value = "����7"
    ws.Range("P3").value = "����8"
    ws.Range("Q3").value = "����9"
    ws.Range("R3").value = "����10"
    ws.Range("S3").value = "����11"
    ws.Range("T3").value = "����12"
    ws.Range("U3").value = "����13"
    ws.Range("V3").value = "����14"
    ws.Range("W3").value = "����15"
    
    
    Const START_ROW = 4
    Dim i As Long
    Dim j As Long
    Dim offset_row As Long: offset_row = 0
    Dim result As GrepResultInfoStruct
    Dim row As Long: row = START_ROW
    Dim params() As String
    
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
        Common.WriteLog "i=" & i
    
        result = results(i)
        
        ws.Cells(row + i, 1).value = result.ResultRaw
        ws.Cells(row + i, 2).value = result.FilePath
        ws.Cells(row + i, 3).value = result.Contents
        ws.Cells(row + i, 4).value = result.ErrorInfo
        
        ws.Cells(row + i, 6).value = result.MethodInfo.Raw
        ws.Cells(row + i, 7).value = result.MethodInfo.Name
        ws.Cells(row + i, 8).value = result.MethodInfo.Ret

        '�������X�g
        params = result.MethodInfo.params
        
        If Common.IsEmptyArray(params) Then
            Common.WriteLog "params is empty."
            GoTo CONTINUE_I
        End If
        
        For j = 0 To UBound(params)
            Common.WriteLog "j=" & j
            ws.Cells(row + i, 9 + j).value = params(j)
        Next j
        
CONTINUE_I:
        
    Next i
    
    Common.WriteLog "OutputSheet E"
End Sub
