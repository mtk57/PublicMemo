Attribute VB_Name = "Module1"
Option Explicit


'�Z���ʒu --------------------------

'[main]�V�[�g
Const OUT_DIR_PATH = "B5"
Const OUT_SQL_TYPE = "B8"

'[�Ώۃe�[�u��]�V�[�g
Const START_TARGET_SHEET_NAME = "B4"

'[XXX�e�[�u��]�V�[�g
Const TABLE_NAME = "B1"
Const START_CLM_NAME = "B3"
Const START_TARGET_CLM = "B4"
Const START_CLM_VALUE = "B5"

'�V�[�g�� --------------------------
Const SHEET_MAIN = "main"
Const SHEET_TARGET_TABLE = "�Ώۃe�[�u��"

'���̑��萔 --------------------------
Const MAX_CLMS = 1000
Const SQL_TYPE_SELECT = "SELECT"
Const SQL_TYPE_UPDATE = "UPDATE"
Const SQL_TYPE_DELETE = "DELETE"

Const IS_TARGET = "��"



Sub �{�^��_Click()
    
On Error GoTo Exception
    
    Dim i, row, clm As Integer
    
    Dim outDirPath As String
    Dim sqlType As String
    Dim fileName As String
    Dim outputFilePath As String
    Dim targetSheetName As String
    Dim cellValue As String
    
    Dim tableName As String
    Dim clmName As String
    Dim isTarget As String
    Dim clmValue As String
    
    Dim valRow, valClm As Integer
    Dim tblRow, tblClm As Integer
    Dim isTargetRow, isTargetClm As Integer
    Dim targetCnt As Integer
    Dim recordCnt As Integer
    
    Dim objRng As Range
    
    Dim targetClmAry() As Integer
    Dim whereCondAry() As String
    Dim sqlAry() As String
    
    
    '�X�N���[���̍X�V�𖳌���
    Application.ScreenUpdating = False
    

    Worksheets(SHEET_MAIN).Select
    
    With ActiveSheet
        outDirPath = Range(OUT_DIR_PATH).Value
        sqlType = Range(OUT_SQL_TYPE).Value
    End With
    
    '---------------------------------------------------------------
    If outDirPath = "" Then
        Err.Raise Number:=12340, _
        Description:="�o�͐�t�H���_�p�X���w�肳��Ă��܂���!"
    End If
    
    If sqlType = "" Then
        Err.Raise Number:=12341, _
        Description:="�o��SQL���w�肳��Ă��܂���!"
    End If
    
    
    
    Worksheets(SHEET_TARGET_TABLE).Select
    
    With ActiveSheet
    
        row = Range(START_TARGET_SHEET_NAME).row
        clm = Range(START_TARGET_SHEET_NAME).Column
        
        '�Ώۃe�[�u���������[�v
        Do
            Worksheets(SHEET_TARGET_TABLE).Select
        
            targetSheetName = Cells(row, clm).Value
            
            If targetSheetName = "" Then
                Exit Do
            End If
            
            If IsExistSheet(targetSheetName) = False Then
                Err.Raise Number:=12342, _
                Description:="�V�[�g��������܂���! (" & targetSheetName & ")"
            End If
            
            Worksheets(targetSheetName).Select
            
            tableName = Range(TABLE_NAME).Value
            
            
            '---------------------------------------------------------------
            tblRow = Range(START_CLM_NAME).row
            tblClm = Range(START_CLM_NAME).Column
            isTargetRow = Range(START_TARGET_CLM).row
            isTargetClm = Range(START_TARGET_CLM).Column
            
            targetCnt = 0
            recordCnt = 0
            Erase targetClmAry
            Erase sqlAry

            '�ΏۃJ���������W
            Do
                cellValue = Cells(tblRow, tblClm).Value
                If cellValue = "" Then
                    '�����񖼂���`����Ă��Ȃ��̂Ŏ��W�I��
                    Exit Do
                End If
                
                isTarget = Cells(isTargetRow, isTargetClm).Value
                If isTarget <> IS_TARGET Then
                    '�Ώۗ�ł͂Ȃ��̂Ŗ���
                    GoTo CONTINUE_CLM
                End If
                
                '�Ώۗ�𔭌�
                ReDim Preserve targetClmAry(targetCnt)
                targetClmAry(targetCnt) = isTargetClm
                targetCnt = targetCnt + 1
                
CONTINUE_CLM:
                tblClm = tblClm + 1
                isTargetClm = isTargetClm + 1
            Loop
            
            
            '---------------------------------------------------------------
            valRow = Range(START_CLM_VALUE).row
            valClm = Range(START_CLM_VALUE).Column
            
            '�Ώۃ��R�[�h�������[�v
            Do
                '��s�`�F�b�N
                Set objRng = Range(Cells(valRow, valClm), Cells(valRow, MAX_CLMS))
                If WorksheetFunction.CountBlank(objRng) = objRng.Count Then
                    '��s�����m�����̂�SQL���o�͂��Ď��̃e�[�u����
                    GoTo OUT_SQL_FILE
                End If
                
                ReDim Preserve sqlAry(recordCnt)

                tblClm = Range(START_CLM_NAME).Column
                
                '�ΏۃJ�����������[�v���āAWHERE����쐬
                ReDim Preserve whereCondAry(targetCnt - 1)
                For i = LBound(targetClmAry) To UBound(targetClmAry)
                    clmName = Cells(tblRow, targetClmAry(i)).Value
                    clmValue = Cells(valRow, targetClmAry(i)).Value
                    whereCondAry(i) = clmName & " = '" & clmValue & "'"
                Next i
                
                'SQL�쐬
                Select Case sqlType
                    Case SQL_TYPE_SELECT
                        sqlAry(recordCnt) = "select * from " & tableName & " where " & Join(whereCondAry, " and ") & ";"
                    Case SQL_TYPE_DELETE
                        sqlAry(recordCnt) = "delete from " & tableName & " where " & Join(whereCondAry, " and ") & ";"
                    Case Else
                        GoTo CONTINUE
                End Select
                
                recordCnt = recordCnt + 1
                valRow = valRow + 1
                
            Loop    '�Ώۃ��R�[�h�������[�v
            
OUT_SQL_FILE:
            '-----------------------------------------------------------
            'SQL���t�@�C���o��
            Select Case sqlType
                Case SQL_TYPE_SELECT
                    fileName = "select.sql"
                Case SQL_TYPE_DELETE
                    fileName = "delete.sql"
                Case Else
                    GoTo CONTINUE
            End Select
            
            outputFilePath = outDirPath & "\" & fileName
            Call AppendOutputFile(outputFilePath, sqlAry)

CONTINUE:
            row = row + 1
               
        Loop    '�Ώۃe�[�u���������[�v
        
    End With
    
    
FINISH:
    Application.ScreenUpdating = False
    MsgBox "Success!"
    Worksheets(SHEET_MAIN).Select
    Exit Sub
    
Exception:
    Application.ScreenUpdating = False
    MsgBox Err.Number & vbCrLf & Err.Description
    Worksheets(SHEET_MAIN).Select
End Sub

Sub CreateOutputFile(ByVal filePath As String)
    If IsExistOutputFile(filePath) = True Then
        Exit Sub
    End If

    Call CreateObject("Scripting.FileSystemObject").CreateTextFile(filePath)
End Sub

Sub AppendOutputFile(ByVal filePath As String, ByRef data() As String)
    If IsExistOutputFile(filePath) = False Then
        Call CreateOutputFile(filePath)
    End If

    Open filePath For Append As #2
    
    Dim i As Integer
    For i = LBound(data) To UBound(data)
        Print #2, data(i)
    Next i
    
    Close #2
End Sub

Function IsExistOutputFile(ByVal filePath As String) As Boolean
    With CreateObject("Scripting.FileSystemObject")
        If .FileExists(filePath) Then
            IsExistOutputFile = True
        Else
            IsExistOutputFile = False
        End If
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
