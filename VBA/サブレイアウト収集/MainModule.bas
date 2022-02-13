Attribute VB_Name = "MainModule"
'�萔
Const KEY_FILE_PATH = "FILE_PATH"
Const KEY_INPUT_SHEET_NAME = "INPUT_SHEET_NAME"
Const KEY_SUBLAYOUT_CLM = "SUBLAYOUT_CLM"
Const KEY_SUBLAYOUT_NAME_CELL_POS = "SUBLAYOUT_NAME_CELL_POS"
Const KEY_COLLECT_START_ROW = "COLLECT_START_ROW"
Const KEY_STOPPER_CLM = "STOPPER_CLM"
Const DICT = "Scripting.Dictionary"

Const MAX_ROWS = 10000

Sub �{�^��1_Click()

On Error GoTo Exception
        
    Set map = CreateObject(DICT)
    
    Worksheets("main").Select

    '�c�[���ɕK�v�ȏ��̓}�b�v�ŊǗ�����
    map.Add KEY_FILE_PATH, Range("B5").Value
    map.Add KEY_INPUT_SHEET_NAME, Range("B9").Value
    map.Add KEY_SUBLAYOUT_CLM, Range("B12").Value
    map.Add KEY_SUBLAYOUT_NAME_CELL_POS, Range("B15").Value
    map.Add KEY_COLLECT_START_ROW, Range("B18").Value
    map.Add KEY_STOPPER_CLM, Range("B21").Value

    '�{�c�[���̃V�[�g��ΏۂƂ��邩�A�w�E�t�@�C���̃V�[�g��ΏۂƂ��邩�̕��򏈗�
    If map(KEY_FILE_PATH) = "" Then
        '�{�c�[���̃V�[�g��Ώ�
        If IsExistSheet(map(KEY_INPUT_SHEET_NAME)) = False Then
            MsgBox "���̓V�[�g�����݂���܂���"
            Exit Sub
        End If
        
        Main (map)
        
        Worksheets("main").Select
        
    Else
        '�w�E�t�@�C���̃V�[�g��Ώ�
        Application.DisplayAlerts = False
        Workbooks.Open map(KEY_FILE_PATH)
        Application.DisplayAlerts = True
        
        If IsExistSheet(map(KEY_INPUT_SHEET_NAME)) = False Then
            MsgBox "���̓V�[�g�����݂���܂���"
            Exit Sub
        End If
        
        Main (map)

    End If

    ShowInfoMsgBox "�I���܂���"
    
    Exit Sub

Exception:
    MsgBox Err.Number & vbCrLf & Err.Description

End Sub

'���C������
Function Main(ByVal map As Object)
    Dim ret As String
    
    Dim in_sheet, out_sheet As String
    Dim sublayoutName, sublayoutSheetName As String
    
    '���C�����C�A�E�g��
    in_sheet = map(KEY_INPUT_SHEET_NAME)

    '���W�J�n
    ret = CollectSublayout(map, in_sheet)
    
    '�߂�l����ȊO�̓G���[�Ƃ���
    If ret <> "" Then
        Call ShowErrMsgBox(ret)
    End If
    
    Exit Function

End Function


'���W�������C��
Function CollectSublayout(ByVal map As Object, ByVal sheetName As String) As String
    Dim i  As Integer
    Dim collect_start_row, sublayoutCount As Integer
    Dim ret, stopperclm, sublayoutclm, beforeSheetName As String
    Dim copy_rows As String
    Dim offset As Integer

    ret = ""
    CollectSublayout = ""

    collect_start_row = val(map(KEY_COLLECT_START_ROW))
    sublayoutclm = map(KEY_SUBLAYOUT_CLM)
    stopperclm = map(KEY_STOPPER_CLM)

    Worksheets(sheetName).Select

    With ActiveSheet
        '�T�u���C�A�E�g������`����Ă����𒲂ׂāA�V�[�g���̃T�u���C�A�E�g�������߂�
        sublayoutCount = CountSublayout(map, sheetName)
        
        If sublayoutCount = 0 Then
            '�T�u���C�A�E�g��1���Ȃ��ꍇ�͐���I��
            Exit Function
        End If
        
        '���W�J�n�s�����s���m�܂Ń��[�v
        For i = collect_start_row To MAX_ROWS
            
            If IsEmpty(Cells(i, stopperclm).Value) Then
                '��s�����m�����̂Ő���I��
                CollectSublayout = ""
                Exit Function
            End If
            
            
            '�T�u���C�A�E�g���̒�`�񂩂�l���擾
            sublayoutName = Cells(i, sublayoutclm).Value

            
            If sublayoutName = "" Then
                '�T�u���C�A�E�g��������`�Ȃ̂Ŗ���
                GoTo CONTINUE_FOR
            End If
            
            
            '�T�u���C�A�E�g���𔭌������̂ŁA�Ή�����V�[�g����������
            sublayoutSheetName = FindSheetName(map, sublayoutName)
            
            If sublayoutSheetName = "" Then
                '�Ή�����V�[�g��������Ȃ������̂ŃG���[�Ƃ���
                CollectSublayout = "�Ή�����T�u���C�A�E�g�̃V�[�g�����݂��܂���(" & sublayoutName & ")"
                Exit Function
            End If
            
            '���݂̃V�[�g����ޔ�
            beforeSheetName = ActiveSheet.Name
            
            '�T�u���C�A�E�g���̃V�[�g������W�i�ċA�Ăяo���j
            ret = CollectSublayout(map, sublayoutSheetName)
            
            '�ޔ����Ă����V�[�g����I��
            Worksheets(beforeSheetName).Select
            
            '�߂�l����ȊO�̓G���[�Ƃ���
            If ret <> "" Then
                CollectSublayout = ret
                Exit Function
            End If
            
            
            '�T�u���C�A�E�g�̓��e���R�s�[���đ}��
            
            '�R�s�[�Ώۂ̍s�͈͂��擾
            copy_rows = GetCopyRows(map, sublayoutSheetName)
            
            Worksheets(beforeSheetName).Select
            
            '�}����̍s�ʒu�̂��߂̃I�t�Z�b�g���擾����
            offset = Worksheets(sublayoutSheetName).Range(copy_rows).Rows.count
            
            '�T�u���C�A�E�g�̓��e���R�s�[
            Worksheets(sublayoutSheetName).Range(copy_rows).Copy
            
            '�T�u���C�A�E�g�̓��e��}��
            Worksheets(beforeSheetName).Rows(i + 1).Insert Shift:=xlDown
            
            '�J�����g�s�ʒu���X�V
            i = i + offset
            
            
CONTINUE_FOR:
        Next i
    
    End With
    

End Function

'�R�s�[�Ώۂ̍s�͈͂��擾
Function GetCopyRows(ByVal map As Object, ByVal sheetName As String) As String
    Dim i, collect_start_row, end_row As Integer
    Dim stopperclm As String
    Dim ret As String

    ret = ""
    collect_start_row = val(map(KEY_COLLECT_START_ROW))
    end_row = 0
    stopperclm = map(KEY_STOPPER_CLM)
    
    Worksheets(sheetName).Select
    
    With ActiveSheet
        '���W�J�n�s�����s���m�܂Ń��[�v
        For i = collect_start_row To MAX_ROWS
            If IsEmpty(Cells(i, stopperclm).Value) Then
                '��s�����m
                
                '�R�s�[����s�̏I���ʒu
                end_row = i - 1
                
                '�R�s�[����s�͈̔͂𕶎���Ŏ擾
                ret = Range(collect_start_row & ":" & end_row).Address
                
                GetCopyRows = ret
                
                Exit Function
            End If
        Next i
    End With
    
End Function

'�T�u���C�A�E�g������`����Ă����𒲂ׂāA�V�[�g���̃T�u���C�A�E�g�������߂�
Function CountSublayout(ByVal map As Object, ByVal sheetName As String) As Integer
    Dim ret, collect_start_row As Integer
    Dim sublayoutName, sublayoutclm, stopperclm As String

    ret = 0
    collect_start_row = val(map(KEY_COLLECT_START_ROW))
    sublayoutclm = map(KEY_SUBLAYOUT_CLM)
    stopperclm = map(KEY_STOPPER_CLM)
    
    Worksheets(sheetName).Select

    With ActiveSheet
        '���W�J�n�s�����s���m�܂Ń��[�v
        For i = collect_start_row To MAX_ROWS
            If IsEmpty(Cells(i, stopperclm).Value) Then
                '��s�����m�����̂ŏI��
                CountSublayout = ret
                Exit Function
            End If
            
            '�T�u���C�A�E�g���̒�`�񂩂�l���擾
            sublayoutName = Cells(i, sublayoutclm).Value
            
            If sublayoutName = "" Then
                '�T�u���C�A�E�g��������`�Ȃ̂Ŗ���
                GoTo CONTINUE_FOR
            End If
            
            '���������������X�V
            ret = ret + 1
                        
CONTINUE_FOR:
        Next i
    
    End With
    
    CountSublayout = ret

End Function

'�T�u���C�A�E�g���ƑΉ�����V�[�g����������
Function FindSheetName(ByVal map As Object, ByVal sublayoutName As String) As String
    Dim sublayoutname_pos As String
    
    sublayoutname_pos = map(KEY_SUBLAYOUT_NAME_CELL_POS)

    '�S�V�[�g������
    For Each ws In Worksheets
        If ws.Range(sublayoutname_pos).Value = sublayoutName Then
            FindSheetName = ws.Name
            Exit Function
        End If
    Next ws
    
    FindSheetName = ""

End Function

'�w�肳�ꂽ�V�[�g�����݂��邩��Ԃ�
Function IsExistSheet(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.Name = sheetName Then
            IsExistSheet = True
            Exit Function
        End If
    Next ws
    
    IsExistSheet = False
End Function

'���b�Z�[�W�{�b�N�X�i���j
Sub ShowInfoMsgBox(msg As String)
    MsgBox msg, vbInformation, ThisWorkbook.Name
End Sub

'���b�Z�[�W�{�b�N�X�i�I�}�[�N�j
Sub ShowErrMsgBox(msg As String)
    MsgBox msg, vbExclamation, ThisWorkbook.Name
End Sub

