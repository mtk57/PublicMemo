Attribute VB_Name = "MainModule"
'�萔
Const MAIN_SHEET = "main"
Const KEY_FILE_PATH = "FILE_PATH"
Const KEY_INPUT_SHEET_NAME = "INPUT_SHEET_NAME"
Const KEY_OUTPUT_SHEET_NAME = "OUTPUT_SHEET_NAME"
Const DICT = "Scripting.Dictionary"
Const XMLDOC = "MSXML2.DOMDocument.6.0"

'XML�^�O
Const TAG_ROOT = "datamodel"
Const TAG_DAO = "dao"
Const TAG_TABLE = "table"
Const TAG_RECORD = "record"
Const TAG_BEFORE = "before"
Const TAG_AFTER = "after"

Const DELI = ","
Const DELI2 = "="

Sub �{�^��1_Click()

On Error GoTo Exception
        
    Set map = CreateObject(DICT)
    
    Worksheets(MAIN_SHEET).Select

    map.Add KEY_FILE_PATH, Range("B5").Value
    map.Add KEY_INPUT_SHEET_NAME, Range("B9").Value
    map.Add KEY_OUTPUT_SHEET_NAME, Range("B11").Value


    If map(KEY_FILE_PATH) = "" Then
        'XML�t�@�C���p�X���w�肵�Ă��Ȃ��ꍇ
        
        If IsExistSheet(map(KEY_INPUT_SHEET_NAME)) = False Then
            MsgBox "���̓V�[�g�����݂���܂���"
            Exit Sub
        End If
        
        '�����ŃV�[�g�̓��e���ꎞ�t�@�C���ɕۑ�����
        'TODO
        
        '�ۑ������p�X�Ńt�@�C���p�X���X�V����
        'TODO
        
        Main (map)
    Else
        'XML�t�@�C���p�X���w�肵���ꍇ
        Main (map)

    End If

    Worksheets(MAIN_SHEET).Select

    MsgBox "�I���܂���"
    
    Exit Sub

Exception:
    MsgBox Err.Number & vbCrLf & Err.Description

End Sub

Function Main(ByVal map As Object)
    Dim i, j As Integer
    Dim out_sheet As String
    Dim objXmlDoc As Object
    Dim objNodeRoot As Object
    Dim objNodeDao As Object
    Dim objNodeTable As Object
    Dim objNodeRecord As Object
    Dim objNodeRecordChild As Object
    Dim objAttr As Object
    Dim row As Integer
    Dim dao_name As String
    Dim table_name As String
    Dim record As String
    Dim varRecord As Variant
    Dim varRecord2 As Variant
    Dim ws As Worksheet
    
    Set objXmlDoc = CreateObject(XMLDOC)

    out_sheet = map(KEY_OUTPUT_SHEET_NAME)
    row = 1

    AddSheet (out_sheet)
    
    Set ws = Worksheets(out_sheet)

    'XML�ǂݍ���
    objXmlDoc.Load (map(KEY_FILE_PATH))
    
    If objXmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox objXmlDoc.parseError.reason, vbCritical
        Exit Function
    End If
    
    '���[�g�̃^�O�����邩?
    Set objNodeRoot = objXmlDoc.SelectSingleNode("//" & TAG_ROOT)
    
    If objNodeRoot Is Nothing Then
        MsgBox "���[�g�^�O������܂���B(" & TAG_ROOT & ")"
        Exit Function
    End If
    
    '�q�^�O�����`�F�b�N
    If objNodeRoot.ChildNodes.Length = 0 Then
        MsgBox "���[�g�^�O�̎q�^�O������܂���B(" & TAG_DAO & ")"
        Exit Function
    End If
    
    'DAO�^�O�̃��X�g�������[�v
    For Each objNodeDao In objNodeRoot.ChildNodes
        If objNodeDao.nodename <> TAG_DAO Then
            GoTo CONTINUE_DAO
        End If
        
        '��������T��
        dao_name = ""
        For Each objAttr In objNodeDao.Attributes
            If objAttr.Name = "id" Then
                dao_name = objAttr.Value
                Exit For
            End If
        Next objAttr
        
        Set objAttr = Nothing
        
        If dao_name = "" Then
            MsgBox "DAO�^�O��id����������܂���B"
            Exit Function
        End If
        
        ws.Cells(row, 1).Value = dao_name
        row = row + 1
        
        'TABLE�^�O�̃��X�g�������[�v
        For Each objNodeTable In objNodeDao.ChildNodes
            If objNodeTable.nodename <> TAG_TABLE Then
                GoTo CONTINUE_TABLE
            End If
            
            '��������T��
            table_name = ""
            For Each objAttr In objNodeTable.Attributes
                If objAttr.Name = "id" Then
                    table_name = objAttr.Value
                    Exit For
                End If
            Next objAttr
            
            Set objAttr = Nothing
            
            If table_name = "" Then
                MsgBox "TABLE�^�O��id����������܂���B"
                Exit Function
            End If
            
            ws.Cells(row, 2).Value = table_name
            row = row + 1
            
            'RECORD�^�O�̃��X�g�������[�v
            For Each objNodeRecord In objNodeTable.ChildNodes
                If objNodeRecord.nodename <> TAG_RECORD Then
                    GoTo CONTINUE_RECORD
                End If
                
                'RECORD�^�O�̎q�^�O�̃��X�g�������[�v
                For Each objNodeRecordChild In objNodeRecord.ChildNodes
                    If objNodeRecordChild.nodename <> TAG_BEFORE And objNodeRecordChild.nodename <> TAG_AFTER Then
                        GoTo CONTINUE_RECORD_CHILD
                    End If
                    
                    record = objNodeRecordChild.Text
                    
                    varRecord = Split(record, DELI)
                    
                    '�o��
                    ws.Cells(row, 3).Value = objNodeRecordChild.nodename
                    row = row + 1
                    
                    For i = 0 To UBound(varRecord)
                        varRecord2 = Split(varRecord(i), DELI2)
                        
                        ws.Cells(row, 4 + i).Value = varRecord2(0)
                        ws.Cells(row + 1, 4 + i).Value = varRecord2(1)
                    
                    Next i
                    
                    row = row + 1
                
CONTINUE_RECORD_CHILD:
                Next objNodeRecordChild
            
CONTINUE_RECORD:
            Next objNodeRecord
            
CONTINUE_TABLE:
        Next objNodeTable
        
CONTINUE_DAO:
    Next objNodeDao
    
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


