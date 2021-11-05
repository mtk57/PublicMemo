Attribute VB_Name = "MainModule"
'�萔
Const KEY_FILE_PATH = "FILE_PATH"
Const KEY_INPUT_SHEET_NAME = "INPUT_SHEET_NAME"
Const KEY_OUTPUT_SHEET_NAME = "OUTPUT_SHEET_NAME"
Const KEY_MAX_ROWS = "MAX_ROWS"
Const KEY_MAX_LEVELS = "MAX_LEVELS"
Const KEY_LEVEL = "LEVEL"
Const KEY_ITEM = "ITEM"
Const KEY_TYPE = "TYPE"
Const KEY_SIZE = "SIZE"
Const DICT = "Scripting.Dictionary"

Sub �{�^��1_Click()

On Error GoTo Exception
        
    Set map = CreateObject(DICT)
    
    Worksheets("main").Select

    map.Add KEY_FILE_PATH, Range("B5").Value
    map.Add KEY_INPUT_SHEET_NAME, Range("B9").Value
    map.Add KEY_OUTPUT_SHEET_NAME, Range("B11").Value
    map.Add KEY_MAX_ROWS, Range("B14").Value
    map.Add KEY_MAX_LEVELS, Range("B17").Value
    map.Add KEY_LEVEL, Range("J5").Value
    map.Add KEY_ITEM, Range("J6").Value
    map.Add KEY_TYPE, Range("J7").Value
    map.Add KEY_SIZE, Range("J8").Value
    

    If map(KEY_FILE_PATH) = "" Then
        If IsExistSheet(map(KEY_INPUT_SHEET_NAME)) = False Then
            MsgBox "���̓V�[�g�����݂���܂���"
            Exit Sub
        End If
        
        Main (map)
        
        Worksheets("main").Select
        
    Else
        Application.DisplayAlerts = False
        Workbooks.Open map(KEY_FILE_PATH)
        Application.DisplayAlerts = True
        
        If IsExistSheet(map(KEY_INPUT_SHEET_NAME)) = False Then
            MsgBox "���̓V�[�g�����݂���܂���"
            Exit Sub
        End If
        
        Main (map)

    End If

    MsgBox "�I���܂���"
    
    Exit Sub

Exception:
    MsgBox Err.Number & vbCrLf & Err.Description

End Sub

Function Main(ByVal map As Object)
    Dim i, j, val_wk, row, out_row As Integer
    Dim wk, wk2 As String
    
    Dim max_rows, max_levels As Integer
    Dim lv, next_lv As Integer
    Dim in_sheet, out_sheet As String
    
    Dim lv_clm, item_clm, type_clm, size_clm As Integer
    Dim level_pos, item_pos, type_pos, size_pos As String
    
    Dim ary_level() As Variant
    Dim ary_item() As Variant
    Dim ary_wk As Variant
    
    '�o�͗p�}�b�v�i�q�������Ȃ����ڗp�j
    Set out_map = CreateObject(DICT)

    '�X�^�b�N�i�q�������ڗp�j
    Set s_level = New Stack
    Set s_item = New Stack

    out_row = 1

    max_rows = val(map(KEY_MAX_ROWS))
    max_levels = val(map(KEY_MAX_LEVELS)) + 1
    
    level_pos = map(KEY_LEVEL)
    item_pos = map(KEY_ITEM)
    type_pos = map(KEY_TYPE)
    size_pos = map(KEY_SIZE)
    
    in_sheet = map(KEY_INPUT_SHEET_NAME)
    out_sheet = map(KEY_OUTPUT_SHEET_NAME)

    AddSheet (out_sheet)

    Worksheets(in_sheet).Select
    
    With ActiveSheet
        row = .Range(level_pos).row
        lv_clm = .Range(level_pos).Column
        item_clm = .Range(item_pos).Column
        type_clm = Range(type_pos).Column
        size_clm = Range(size_pos).Column
    
        For i = 1 To max_rows
            wk = Cells(row + i, lv_clm).Value
            '���s�̒l
            wk2 = Cells(row + i + 1, lv_clm).Value
            
            If wk = "" Or IsNumeric(wk) = False Then
                GoTo CONTINUE_FOR
            End If
            
            lv = val(wk)
            
            If wk2 = "" Or IsNumeric(wk2) = False Or val(wk2) = lv Or val(wk2) < lv Then
                '�q�������Ȃ����ڂł��邱�Ƃ��m��
                
                'Level���������Ȃ����ꍇ�A�X�^�b�N����폜����
                If s_level.count > lv Then
                    For j = 0 To s_level.count - lv
                        s_level.pop
                        s_item.pop
                    Next j
                End If
                
                '--------------------------
                '�o�͗p�}�b�v�ɒl��ݒ�
                '--------------------------
                
                '�q�������ڂ̓X�^�b�N�̒l���g�p���Đݒ�
                ary_level = s_level.getContents
                ary_item = s_item.getContents
                For j = 1 To s_item.count
                    out_map.Add ary_level(j), ary_item(j)
                Next j
                
                '�q�������Ȃ����ڂ̐ݒ�
                out_map.Add lv, Cells(row + i, item_clm).Value
                out_map.Add max_levels, Cells(row + i, type_clm).Value
                out_map.Add max_levels + 1, Cells(row + i, size_clm).Value
                
                'for DEBUG
                'ary_wk = out_map.items
                
                '�V�[�g�ɏo��
                Worksheets(out_sheet).Select
                For Each key_clm In out_map
                    Cells(out_row, key_clm).Value = out_map(key_clm)
                Next
                
                '���̏o�͍s���X�V
                out_row = out_row + 1
                
                out_map.RemoveAll
                
                Worksheets(in_sheet).Select
                
            Else
                '�q��������

                'Level��Item���X�^�b�N�ɐς�
                s_level.push lv
                s_item.push Cells(row + i, item_clm).Value
            End If
            
            
CONTINUE_FOR:
        Next i
    
    End With

End Function

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

Function AddSheet(ByVal sheetName As String)
    If IsExistSheet(sheetName) = True Then
        Application.DisplayAlerts = False
        Sheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If

    Worksheets.Add.Name = sheetName
End Function


