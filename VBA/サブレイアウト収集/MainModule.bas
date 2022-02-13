Attribute VB_Name = "MainModule"
'定数
Const KEY_FILE_PATH = "FILE_PATH"
Const KEY_INPUT_SHEET_NAME = "INPUT_SHEET_NAME"
Const KEY_SUBLAYOUT_CLM = "SUBLAYOUT_CLM"
Const KEY_SUBLAYOUT_NAME_CELL_POS = "SUBLAYOUT_NAME_CELL_POS"
Const KEY_COLLECT_START_ROW = "COLLECT_START_ROW"
Const KEY_STOPPER_CLM = "STOPPER_CLM"
Const DICT = "Scripting.Dictionary"

Const MAX_ROWS = 10000

Sub ボタン1_Click()

On Error GoTo Exception
        
    Set map = CreateObject(DICT)
    
    Worksheets("main").Select

    'ツールに必要な情報はマップで管理する
    map.Add KEY_FILE_PATH, Range("B5").Value
    map.Add KEY_INPUT_SHEET_NAME, Range("B9").Value
    map.Add KEY_SUBLAYOUT_CLM, Range("B12").Value
    map.Add KEY_SUBLAYOUT_NAME_CELL_POS, Range("B15").Value
    map.Add KEY_COLLECT_START_ROW, Range("B18").Value
    map.Add KEY_STOPPER_CLM, Range("B21").Value

    '本ツールのシートを対象とするか、指摘ファイルのシートを対象とするかの分岐処理
    If map(KEY_FILE_PATH) = "" Then
        '本ツールのシートを対象
        If IsExistSheet(map(KEY_INPUT_SHEET_NAME)) = False Then
            MsgBox "入力シート名がみつかりません"
            Exit Sub
        End If
        
        Main (map)
        
        Worksheets("main").Select
        
    Else
        '指摘ファイルのシートを対象
        Application.DisplayAlerts = False
        Workbooks.Open map(KEY_FILE_PATH)
        Application.DisplayAlerts = True
        
        If IsExistSheet(map(KEY_INPUT_SHEET_NAME)) = False Then
            MsgBox "入力シート名がみつかりません"
            Exit Sub
        End If
        
        Main (map)

    End If

    ShowInfoMsgBox "終わりました"
    
    Exit Sub

Exception:
    MsgBox Err.Number & vbCrLf & Err.Description

End Sub

'メイン処理
Function Main(ByVal map As Object)
    Dim ret As String
    
    Dim in_sheet, out_sheet As String
    Dim sublayoutName, sublayoutSheetName As String
    
    'メインレイアウト名
    in_sheet = map(KEY_INPUT_SHEET_NAME)

    '収集開始
    ret = CollectSublayout(map, in_sheet)
    
    '戻り値が空以外はエラーとする
    If ret <> "" Then
        Call ShowErrMsgBox(ret)
    End If
    
    Exit Function

End Function


'収集処理メイン
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
        'サブレイアウト名が定義されている列を調べて、シート内のサブレイアウト数を求める
        sublayoutCount = CountSublayout(map, sheetName)
        
        If sublayoutCount = 0 Then
            'サブレイアウトが1つもない場合は正常終了
            Exit Function
        End If
        
        '収集開始行から空行検知までループ
        For i = collect_start_row To MAX_ROWS
            
            If IsEmpty(Cells(i, stopperclm).Value) Then
                '空行を検知したので正常終了
                CollectSublayout = ""
                Exit Function
            End If
            
            
            'サブレイアウト名の定義列から値を取得
            sublayoutName = Cells(i, sublayoutclm).Value

            
            If sublayoutName = "" Then
                'サブレイアウト名が未定義なので無視
                GoTo CONTINUE_FOR
            End If
            
            
            'サブレイアウト名を発見したので、対応するシートを検索する
            sublayoutSheetName = FindSheetName(map, sublayoutName)
            
            If sublayoutSheetName = "" Then
                '対応するシートが見つからなかったのでエラーとする
                CollectSublayout = "対応するサブレイアウトのシートが存在しません(" & sublayoutName & ")"
                Exit Function
            End If
            
            '現在のシート名を退避
            beforeSheetName = ActiveSheet.Name
            
            'サブレイアウト名のシートから収集（再帰呼び出し）
            ret = CollectSublayout(map, sublayoutSheetName)
            
            '退避していたシート名を選択
            Worksheets(beforeSheetName).Select
            
            '戻り値が空以外はエラーとする
            If ret <> "" Then
                CollectSublayout = ret
                Exit Function
            End If
            
            
            'サブレイアウトの内容をコピーして挿入
            
            'コピー対象の行範囲を取得
            copy_rows = GetCopyRows(map, sublayoutSheetName)
            
            Worksheets(beforeSheetName).Select
            
            '挿入後の行位置のためのオフセットを取得する
            offset = Worksheets(sublayoutSheetName).Range(copy_rows).Rows.count
            
            'サブレイアウトの内容をコピー
            Worksheets(sublayoutSheetName).Range(copy_rows).Copy
            
            'サブレイアウトの内容を挿入
            Worksheets(beforeSheetName).Rows(i + 1).Insert Shift:=xlDown
            
            'カレント行位置を更新
            i = i + offset
            
            
CONTINUE_FOR:
        Next i
    
    End With
    

End Function

'コピー対象の行範囲を取得
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
        '収集開始行から空行検知までループ
        For i = collect_start_row To MAX_ROWS
            If IsEmpty(Cells(i, stopperclm).Value) Then
                '空行を検知
                
                'コピーする行の終了位置
                end_row = i - 1
                
                'コピーする行の範囲を文字列で取得
                ret = Range(collect_start_row & ":" & end_row).Address
                
                GetCopyRows = ret
                
                Exit Function
            End If
        Next i
    End With
    
End Function

'サブレイアウト名が定義されている列を調べて、シート内のサブレイアウト数を求める
Function CountSublayout(ByVal map As Object, ByVal sheetName As String) As Integer
    Dim ret, collect_start_row As Integer
    Dim sublayoutName, sublayoutclm, stopperclm As String

    ret = 0
    collect_start_row = val(map(KEY_COLLECT_START_ROW))
    sublayoutclm = map(KEY_SUBLAYOUT_CLM)
    stopperclm = map(KEY_STOPPER_CLM)
    
    Worksheets(sheetName).Select

    With ActiveSheet
        '収集開始行から空行検知までループ
        For i = collect_start_row To MAX_ROWS
            If IsEmpty(Cells(i, stopperclm).Value) Then
                '空行を検知したので終了
                CountSublayout = ret
                Exit Function
            End If
            
            'サブレイアウト名の定義列から値を取得
            sublayoutName = Cells(i, sublayoutclm).Value
            
            If sublayoutName = "" Then
                'サブレイアウト名が未定義なので無視
                GoTo CONTINUE_FOR
            End If
            
            '発見した総数を更新
            ret = ret + 1
                        
CONTINUE_FOR:
        Next i
    
    End With
    
    CountSublayout = ret

End Function

'サブレイアウト名と対応するシートを検索する
Function FindSheetName(ByVal map As Object, ByVal sublayoutName As String) As String
    Dim sublayoutname_pos As String
    
    sublayoutname_pos = map(KEY_SUBLAYOUT_NAME_CELL_POS)

    '全シートを検索
    For Each ws In Worksheets
        If ws.Range(sublayoutname_pos).Value = sublayoutName Then
            FindSheetName = ws.Name
            Exit Function
        End If
    Next ws
    
    FindSheetName = ""

End Function

'指定されたシートが存在するかを返す
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

'メッセージボックス（情報）
Sub ShowInfoMsgBox(msg As String)
    MsgBox msg, vbInformation, ThisWorkbook.Name
End Sub

'メッセージボックス（！マーク）
Sub ShowErrMsgBox(msg As String)
    MsgBox msg, vbExclamation, ThisWorkbook.Name
End Sub

