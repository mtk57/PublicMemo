Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String

'�p�����[�^
Public main_param As MainParam

Private final_sheet_name As String

Private vbp_mapping As RefFiles
Private vbproj_mapping As RefFiles
Private repname_mapping As ReplaceModel

Private final_mapping As ReplaceModel


'--------------------------------------------------------
'���C������
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    '�p�����[�^�̃`�F�b�N�Ǝ��W���s��
    CheckAndCollectParam
    
    'VBP�̃}�b�s���O���쐬����
    CreateVbpMapping
    
    'VBPROJ�̃}�b�s���O���쐬����
    CreateVbprojMapping
    
    '���l�[���}�b�s���O���쐬����
    CreateRenameMapping
    
    '�ŏI�}�b�s���O���쐬����
    CreateFinalMapping
    
    '�ŏI�}�b�s���O���V�[�g�ɏo�͂���
    OutputSheet

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

Private Sub CreateVbpMapping()
    Common.WriteLog "CreateVbpMapping S"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(main_param.GetVbpMappingSheetName())
    
    Dim i As Long
    Dim j As Long
    
    Dim key As String
    Dim val As String
    
    Const START_ROW = 5
    i = 0
    
    Set vbp_mapping = New RefFiles
    
    Do
        key = ws.Cells(START_ROW + i, 1).value
        val = ws.Cells(START_ROW + i, 2).value

        If key = "" Then
            Exit Do
        End If
        
        If val = "" Then
            'ref_file_path����Ȃ͕̂s���ȃ}�b�s���O�f�[�^�Ȃ̂Ŗ�������
            Common.WriteLog "��Bad data found.  key=<" & key & ">, val=<" & val & ">"
            GoTo CONTINUE
        End If
        
        vbp_mapping.SetRowData key, val
        
CONTINUE:
        i = i + 1
    Loop
    
    'for DEBUG
    'For i = 0 To vbp_mapping.GetRowCount()
    '    Dim prj As Variant
    '    Dim paths As Variant
    '
    '    prj = vbp_mapping.GetPrjPathList()(i)
    '    paths = vbp_mapping.GetRefPath(prj)
    '
    '    For j = 0 To UBound(paths)
    '        Common.WriteLog "prj=" & prj & ", path=" & paths(j)
    '    Next j
    'Next i
    
    Common.WriteLog "CreateVbpMapping E"
End Sub

Private Sub CreateVbprojMapping()

    Common.WriteLog "CreateVbprojMapping S"
    
    Dim msg As String: msg = ""
    
    '�O���c�[�����s
    Const MACRO_NAME As String = "Main.Run"
    Dim ret_dict As Variant
        
    Set ret_dict = Application.Run( _
          "'" & _
          main_param.GetVbFileListCreatorPath() & _
          "'!" & _
          MACRO_NAME, _
          main_param.GetVbprojDirPath(), _
          "vbproj", _
          "", _
          "vb", _
          main_param.IsDebug() _
    )

    
    
    '--------------------------------------------------------------------------
    '�O���c�[�������O�ɃR�s�[���Ă���
    '��Dict�N���X�ɂ��̏������������悤�Ƃ��������̂���肭�����Ȃ�����...orz
    '--------------------------------------------------------------------------
    Dim copy_dict As Dict
    Set copy_dict = New Dict
    
    Dim i As Long
    Dim j As Long
    
    Common.WriteLog "GetAllValueCount=" & ret_dict.GetAllValueCount()
    
    For i = 0 To ret_dict.GetKeyCount()
        Dim key As String: key = ret_dict.GetKeys()(i)
        
        Dim values() As String
        values = ret_dict.GetValue(key)
        
        For j = 0 To UBound(values)
            Dim value As String: value = values(j)
            copy_dict.AppendStringArray key, value
            
            Common.WriteLog "key=<" & key & ">, value=<" & value & ">"
        Next j
    Next i
    '--------------------------------------------------------------------------
    
    
    
    '�O���c�[�������
    Common.CloseBook Common.GetFileName(main_param.GetVbFileListCreatorPath()), True
    
    Set vbproj_mapping = New RefFiles
    vbproj_mapping.SetDict copy_dict
    
    Common.WriteLog "CreateVbprojMapping E"

End Sub

Private Sub CreateRenameMapping()
    Common.WriteLog "CreateRenameMapping S"
    
    Const START_ROW = 5
    
    
    '�܂��͍s���𐔂���
    Dim rename_mapping_row_cnt As Long
    Dim i As Long
    Dim src_key As String
    Dim src_val As String
    Dim dst_key As String
    Dim dst_val As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(main_param.GetRenameMappingSheetName())
    
    Do
        src_key = ws.Cells(START_ROW + i, 1).value
        src_val = ws.Cells(START_ROW + i, 2).value
        dst_key = ws.Cells(START_ROW + i, 3).value
        dst_val = ws.Cells(START_ROW + i, 4).value
        
        If src_key = "" And src_val = "" And dst_key = "" And dst_val = "" Then
            '�S�ċ�Ȃ̂ŏI������
            Exit Do
        End If
        
        'If src_key = "" Or src_val = "" Or dst_key = "" Or dst_val = "" Then
        '    '1�ł��󂪂���Εs���ȃ}�b�s���O�f�[�^�Ȃ̂ł����ŏI������
        '    Common.WriteLog "��Bad data found.  src_key=<" & src_key & ">, src_val=<" & src_val & ">, dst_key=<" & dst_key & ">, dst_val=<" & dst_val & ">"
        '    Exit Do
        'End If
       
        i = i + 1
    Loop
    
    If i = 0 Then
        Common.WriteLog "rename_mapping_row_cnt=" & rename_mapping_row_cnt
    End If
    
    rename_mapping_row_cnt = i - 1
    

    'ReplaceModel���쐬����
    Set repname_mapping = New ReplaceModel
    repname_mapping.Init rename_mapping_row_cnt
    
    If rename_mapping_row_cnt < 0 Then
        Common.WriteLog "CreateRenameMapping E-1"
        Exit Sub
    End If
    
    For i = 0 To rename_mapping_row_cnt
        
        src_key = ws.Cells(START_ROW + i, 1).value
        src_val = ws.Cells(START_ROW + i, 2).value
        dst_key = ws.Cells(START_ROW + i, 3).value
        dst_val = ws.Cells(START_ROW + i, 4).value
        
        repname_mapping.Append src_key, src_val, dst_key, dst_val
                
    Next i
    
    Common.WriteLog "CreateRenameMapping E"
End Sub

Private Sub CreateFinalMapping()
    Common.WriteLog "CreateFinalMapping S"

    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim key As String
    Dim values() As String
    Dim value As String
    Dim expect_key As String:
    Dim expect_val As String
    Dim final_key As String
    Dim final_val As String
    Dim ren_key_num As Long
    Dim ren_val_num As Long
    Dim search_key As String

    Set final_mapping = New ReplaceModel
    final_mapping.Init (vbp_mapping.GetAllRefCount())
    
    For i = 0 To vbp_mapping.GetRowCount()
        
        '���҂���L�[�����ۂɎ��W�����L�[�ɑ��݂��邩?
        key = vbp_mapping.GetPrjPathList()(i)
        expect_key = Replace(key, main_param.GetVbpBaseDirPath(), main_param.GetVbprojBaseDirPath())
        expect_key = Replace(expect_key, "." & Common.GetFileExtension(expect_key, True), ".vbproj")
        
        'If vbproj_mapping.IsExistKey(expect_key) = False Then
        '    Common.WriteLog "Project file is not exist. Expect=<" & expect_key & ">"
        '
        '    '�R�Â��ł��Ȃ��������Ƃ�\�����ߍŏI�}�b�s���O�ɓo�^
        '    final_mapping.Append key, value, "Not found.", "Not found."
        '
        '    '���݂��Ȃ��̂Ŗ���
        '    GoTo CONTINUE_I
        'End If
        
        If Common.IsExistsFile(expect_key) = True Then
            'vbproj�t�@�C�������ۂɑ��݂���̂Ŋm��
            final_key = expect_key
        Else
            final_key = ""
            
            '���l�[���}�b�s���O�ɂ��邩�m�F���Ă���΍̗p����
            If repname_mapping.IsEmpty() = False Then
                ren_key_num = repname_mapping.GetIndexSrcPrjPath(key)
            
                If ren_key_num >= 0 Then
                    final_key = repname_mapping.GetDstPrjPath(ren_key_num)
                Else
                    '���l�[���}�b�s���O�ɂ������̂Ńt�@�C���������Ă݂�
                    search_key = Common.SearchFile(main_param.GetVbprojDirPath(), Common.GetFileName(expect_key))
                    
                    If search_key <> "" Then
                        '���������̂Ńq���g�Ƃ��čŏI�}�b�s���O�ɓo�^
                        Common.WriteLog "��VBPROJ PATH NOT FOUND(BUT SEARCH FOUND). vbp=<" & key & ">, vbproj(search)=<" & search_key & ">"
                
                        final_mapping.Append key, value, "��vbproj is search found.<" & search_key & ">", "unknown."
                        
                        GoTo CONTINUE_I
                    End If
                End If
            End If
        End If
        
        If final_key = "" Then
            Common.WriteLog "��VBPROJ PATH NOT FOUND. vbp=<" & key & ">, vbproj(expect)=<" & expect_key & ">"
            
            '�R�Â��ł��Ȃ��������Ƃ�\�����ߍŏI�}�b�s���O�ɓo�^
            final_mapping.Append key, value, "vbproj is not found.", "unknown."
            
            '���݂��Ȃ��̂Ŗ���
            GoTo CONTINUE_I
        End If
        
        'vbproj�͊m��
        Common.WriteLog "��VBPROJ PATH CONFIRM. vbp=<" & key & ">, vbproj=<" & final_key & ">"
        
        
        '���͎Q�ƃt�@�C���Q
        values = vbp_mapping.GetRefPath(key)
        
        For j = 0 To UBound(values)
            
            '���҂���l�����ۂɎ��W�����l�ɑ��݂��邩?
            value = values(j)
            
            expect_val = Replace(value, main_param.GetVbpBaseDirPath(), main_param.GetVbprojBaseDirPath())
            expect_val = Replace(expect_val, "." & Common.GetFileExtension(expect_val, True), ".vb")
            
            If Common.IsExistsFile(expect_val) = True Then
                '���҂���Q�ƃt�@�C�������ۂɑ��݂���̂Ŋm��
                final_val = expect_val
            
                GoTo CONFIRMED
            End If
            
            
            final_val = ""

            '���l�[���}�b�s���O�ɂ��邩�m�F���Ă���΍̗p����
            If repname_mapping.IsEmpty() = False Then
                ren_val_num = repname_mapping.GetIndexSrcRefPath(value)

                If ren_key_num >= 0 Then
                    final_key = repname_mapping.GetDstPrjPath(ren_key_num)
                   GoTo CONFIRMED
                End If
                
                If ren_val_num >= 0 Then
                    final_val = repname_mapping.GetDstRefPath(ren_val_num)
                   GoTo CONFIRMED
                End If
            End If
            
            If final_val = "" Then
                Common.WriteLog "��VBPROJ REF PATH NOT FOUND. vbp ref=<" & value & ">, vbproj ref(expect)=<" & expect_val & ">"
                
                '�R�Â��ł��Ȃ��������Ƃ�\�����ߍŏI�}�b�s���O�ɓo�^
                final_mapping.Append key, value, final_key, "vbproj ref is not found."
                
                '���݂��Ȃ��̂Ŗ���
                GoTo CONTINUE_J
            End If
            
CONFIRMED:
            'vbproj ref�͊m��
            Common.WriteLog "��VBPROJ REF PATH CONFIRM.  vbp ref=<" & value & ">, vbproj ref=<" & final_val & ">"
            
            '�L�[/�l�����ۂɑ��݂���̂ōŏI�}�b�s���O�Ƃ��č̗p
            final_mapping.Append key, value, final_key, final_val
            
CONTINUE_J:
            
        Next j
        
CONTINUE_I:
        
    Next i
    
    Common.WriteLog "CreateFinalMapping E"
End Sub

Private Sub OutputSheet()
    Common.WriteLog "OutputSheet S"
    
    final_sheet_name = Common.GetNowTimeString_OLD()
    
    Common.AddSheet ThisWorkbook, final_sheet_name
    
    '�w�b�_
    Common.UpdateSheet ActiveWorkbook, final_sheet_name, 1, 1, "vbp full path"
    Common.UpdateSheet ActiveWorkbook, final_sheet_name, 1, 2, "vbp ref file full path"
    Common.UpdateSheet ActiveWorkbook, final_sheet_name, 1, 3, "vbproj full path"
    Common.UpdateSheet ActiveWorkbook, final_sheet_name, 1, 4, "vbproj ref file full path"
        
    
    Dim i As Long
    Dim cur_row As Long
    cur_row = 2
    
    For i = 0 To final_mapping.GetRowCount()
        Common.UpdateSheet ActiveWorkbook, final_sheet_name, cur_row, 1, final_mapping.GetSrcPrjPath(i)
        Common.UpdateSheet ActiveWorkbook, final_sheet_name, cur_row, 2, final_mapping.GetSrcRefPath(i)
        Common.UpdateSheet ActiveWorkbook, final_sheet_name, cur_row, 3, final_mapping.GetDstPrjPath(i)
        Common.UpdateSheet ActiveWorkbook, final_sheet_name, cur_row, 4, final_mapping.GetDstRefPath(i)
        cur_row = cur_row + 1
    Next i
    
    Common.WriteLog "OutputSheet E"
End Sub
