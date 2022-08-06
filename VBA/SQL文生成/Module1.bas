Attribute VB_Name = "Module1"
Option Explicit


'セル位置 --------------------------

'[main]シート
Const OUT_DIR_PATH = "B5"
Const OUT_SQL_TYPE = "B8"

'[対象テーブル]シート
Const START_TARGET_SHEET_NAME = "B4"

'[XXXテーブル]シート
Const TABLE_NAME = "B1"
Const START_CLM_NAME = "B3"
Const START_TARGET_CLM = "B4"
Const START_CLM_VALUE = "B5"

'シート名 --------------------------
Const SHEET_MAIN = "main"
Const SHEET_TARGET_TABLE = "対象テーブル"

'その他定数 --------------------------
Const MAX_CLMS = 1000
Const SQL_TYPE_SELECT = "SELECT"
Const SQL_TYPE_UPDATE = "UPDATE"
Const SQL_TYPE_DELETE = "DELETE"

Const IS_TARGET = "○"



Sub ボタン_Click()
    
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
    
    
    'スクリーンの更新を無効化
    Application.ScreenUpdating = False
    

    Worksheets(SHEET_MAIN).Select
    
    With ActiveSheet
        outDirPath = Range(OUT_DIR_PATH).Value
        sqlType = Range(OUT_SQL_TYPE).Value
    End With
    
    '---------------------------------------------------------------
    If outDirPath = "" Then
        Err.Raise Number:=12340, _
        Description:="出力先フォルダパスが指定されていません!"
    End If
    
    If sqlType = "" Then
        Err.Raise Number:=12341, _
        Description:="出力SQLが指定されていません!"
    End If
    
    
    
    Worksheets(SHEET_TARGET_TABLE).Select
    
    With ActiveSheet
    
        row = Range(START_TARGET_SHEET_NAME).row
        clm = Range(START_TARGET_SHEET_NAME).Column
        
        '対象テーブル数分ループ
        Do
            Worksheets(SHEET_TARGET_TABLE).Select
        
            targetSheetName = Cells(row, clm).Value
            
            If targetSheetName = "" Then
                Exit Do
            End If
            
            If IsExistSheet(targetSheetName) = False Then
                Err.Raise Number:=12342, _
                Description:="シートが見つかりません! (" & targetSheetName & ")"
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

            '対象カラムを収集
            Do
                cellValue = Cells(tblRow, tblClm).Value
                If cellValue = "" Then
                    '物理列名が定義されていないので収集終了
                    Exit Do
                End If
                
                isTarget = Cells(isTargetRow, isTargetClm).Value
                If isTarget <> IS_TARGET Then
                    '対象列ではないので無視
                    GoTo CONTINUE_CLM
                End If
                
                '対象列を発見
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
            
            '対象レコード数分ループ
            Do
                '空行チェック
                Set objRng = Range(Cells(valRow, valClm), Cells(valRow, MAX_CLMS))
                If WorksheetFunction.CountBlank(objRng) = objRng.Count Then
                    '空行を検知したのでSQLを出力して次のテーブルへ
                    GoTo OUT_SQL_FILE
                End If
                
                ReDim Preserve sqlAry(recordCnt)

                tblClm = Range(START_CLM_NAME).Column
                
                '対象カラム数分ループして、WHERE句を作成
                ReDim Preserve whereCondAry(targetCnt - 1)
                For i = LBound(targetClmAry) To UBound(targetClmAry)
                    clmName = Cells(tblRow, targetClmAry(i)).Value
                    clmValue = Cells(valRow, targetClmAry(i)).Value
                    whereCondAry(i) = clmName & " = '" & clmValue & "'"
                Next i
                
                'SQL作成
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
                
            Loop    '対象レコード数分ループ
            
OUT_SQL_FILE:
            '-----------------------------------------------------------
            'SQLをファイル出力
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
               
        Loop    '対象テーブル数分ループ
        
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
