Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String

'�p�����[�^
Private main_param As MainParam
Private sub_params() As SubParam
Private form_data As FormData

'--------------------------------------------------------
'���C������
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    CheckAndCollectParam
    
    'frm�t�@�C����ǂݍ���
    LoadFrmFile
    
    '���C�����[�v
    Dim i As Long
    For i = LBound(sub_params) To UBound(sub_params)
        Dim target As SubParam: target = sub_params(i)
        'Common.WriteLog "i=" & i & ":[" & targer_path & "]"
    
    Next i

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
    Const START_ROW = 4
    Dim row As Long: row = START_ROW
    Dim cnt As Long: cnt = 0
    
    Do
        Dim sub_param As SubParam
        Set sub_param = New SubParam
        
        Common.WriteLog "row=" & row
        sub_param.Init row
        sub_param.Validate

        'Common.WriteLog sub_param.GetAllValue()
        
        If sub_param.GetEnable() = "Stopper" Then
            Exit Do
        ElseIf sub_param.GetEnable() = "Disable" Then
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

'frm�t�@�C����ǂݍ���Ńf�[�^���f���ɕϊ�����
Private Sub LoadFrmFile()
    Common.WriteLog "LoadFrmFile S"
    
    Dim contents() As String: contents = GetContents(main_param.GetSrcFrmFilePath())
    
    
    
    Common.WriteLog "LoadFrmFile E"
End Sub

'Designer.vb�t�@�C����ǂݍ���Ńf�[�^���f���ɕϊ�����
Private Sub LoadDesignerVbFile()
    Common.WriteLog "LoadDesignerVbFile S"
    
    Dim contents() As String: contents = GetContents(main_param.GetSrcDesignerVbFilePath())
    
    
    
    Common.WriteLog "LoadDesignerVbFile E"
End Sub

'�t�@�C���̓��e��ǂݍ���
Private Function GetContents(ByVal path As String) As String()
    Common.WriteLog "GetContents S"
    
    Dim raw_contents As String
    
    '�t�@�C���̓��e��ǂݍ���
    If Common.IsSJIS(path) = True Then
        raw_contents = Common.ReadTextFileBySJIS(path)
    ElseIf Common.IsUTF8(path) = True Then
        raw_contents = Common.ReadTextFileByUTF8(path)
    Else
        Err.Raise 53, , "���T�|�[�g�̃G���R�[�h�ł��B(" & path & ")"
    End If
    
    '�t�@�C���̓��e��z��Ɋi�[����
    Dim contents() As String: contents = Split(raw_contents, vbCrLf)
    
    GetContents = Common.DeleteEmptyArray(contents)
    
    Common.WriteLog "GetContents E"
End Function
