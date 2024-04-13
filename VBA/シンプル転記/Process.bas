Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String
Private DQ As String

'パラメータ
Private main_param As MainParam
Private sub_params() As subparam

'メイン処理
Public Sub Run()
    Common.WriteLog "Run S"

    Worksheets("main").Activate
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    'パラメータのチェックと収集を行う
    CheckAndCollectParam
    
    'Sub Paramを順に実行していく
    ExecSubParam
    
    Common.WriteLog "Run E"
End Sub

'パラメータのチェックと収集を行う
Private Sub CheckAndCollectParam()
    Common.WriteLog "CheckAndCollectParam S"
    
    Dim err_msg As String

    'Main Params
    Set main_param = New MainParam
    main_param.Init
    main_param.Validate
    
    Common.WriteLog main_param.GetAllValue()

    'Sub Params
    Const START_ROW = 17
    Dim row As Long: row = START_ROW
    Dim cnt As Long: cnt = 0
    
    Do
        Dim sub_param As subparam
        Set sub_param = New subparam
        
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

'Sub Paramを順に実行していく
Private Sub ExecSubParam()
    Common.WriteLog "ExecSubParam S"
    
    If Common.IsEmptyArray(sub_params) = True Then
        Err.Raise 53, , "有効なSub paramがありません"
    End If

    Dim i As Integer
    Dim copy_datas() As CopyData
    Dim sub_param As subparam
    
    For i = LBound(sub_params) To UBound(sub_params)
        Set sub_param = sub_params(i)
        
        Common.WriteLog "■Main Loop (SubParam Row#=" & sub_param.GetSubParamRowNumber() & ")"
        
        '転記元データを収集する
        copy_datas = CollectSrcDatas(sub_param)
            
        If Common.IsEmptyArray(copy_datas) = True Then
            Common.WriteLog "転記元データがありません。"
            GoTo CONTINUE_FOR
        End If
            
        '転記元データを転記先に転記する
        If sub_param.IsDstMultiCopy() = False Then
            Transcription sub_param, copy_datas
        Else
            TranscriptionForMultiCopy sub_param, copy_datas
        End If

CONTINUE_FOR:

    Next i
    
    Common.WriteLog "ExecSubParam E"
End Sub

'転記元データを収集する
Private Function CollectSrcDatas(ByRef sub_param As subparam) As CopyData()
    Common.WriteLog "CollectSrcDatas S"

    Dim ws As Worksheet
    Dim copy_datas() As CopyData
    Dim copy_data As CopyData
    Dim cnt As Long
    Dim cell As Range

    'SRCファイルパスのSRCシート名を開く
    Const READ_ONLY_FLG = True
    Const VISIBLE_FLG = True
    Set ws = Common.GetSheet( _
                sub_param.GetSrcFilePath(), _
                sub_param.GetSrcSheetName(), _
                READ_ONLY_FLG, _
                VISIBLE_FLG _
             )
    
    'SRC検索列の黄色セルに対応するSRC転記列の値を収集する
    Dim key_rng As Range
    Dim value_rng As Range
    Dim key_clm As String: key_clm = sub_param.GetSrcFindClm()
    Dim val_clm As String: val_clm = sub_param.GetSrcTranClm()
    
    Set key_rng = ws.Range(key_clm & "1:" & key_clm & ws.Cells(ws.Rows.count, key_clm).End(xlUp).row)
    Set value_rng = ws.Range(val_clm & "1:" & val_clm & ws.Cells(ws.Rows.count, val_clm).End(xlUp).row)

    cnt = 0
    For Each cell In key_rng
        '収集対象は黄色セルのみとする
        If cell.Interior.Color = RGB(255, 255, 0) Then
            ReDim Preserve copy_datas(cnt)
            Set copy_data = New CopyData
            copy_data.Init cell.value, value_rng.Cells(cell.row, 1).value
            Set copy_datas(cnt) = copy_data
            cnt = cnt + 1
        End If
    Next cell
    
    If main_param.IsNotClose() = False Then
        'SRCファイルを閉じる
        Common.CloseBook (Common.GetFileName(sub_param.GetSrcFilePath()))
    End If
    
    CollectSrcDatas = copy_datas
    
    Common.WriteLog "CollectSrcDatas E"
End Function

'転記する
Private Sub Transcription(ByRef sub_param As subparam, ByRef copy_datas() As CopyData)
    Common.WriteLog "Transcription S"
    
    Dim ws As Worksheet
    Dim book_name As String
    Dim row As Long
    Dim keyword As String
    Dim found_row As Long
    Dim trans_rng As Range
    Dim copy_data As CopyData
    
    'DSTファイルパスのDSTシート名を開く
    Set ws = Common.GetSheet(sub_param.GetDstFilePath(), sub_param.GetDstSheetName(), False, True)
    book_name = Common.GetFileName(sub_param.GetDstFilePath())
    
    Dim last_row As Long: last_row = Common.GetLastRowFromWorksheet(ws, sub_param.GetDstFindClm())
    
    'SRC検索列の値が、DST検索列にあるか検索する
    'あれば、SRC転記列の値をDST転記列に入れる
    For row = LBound(copy_datas, 1) To UBound(copy_datas, 1)
    
        Set copy_data = copy_datas(row)
        keyword = copy_data.GetKey()
        
        If keyword = "" Then
            GoTo CONTINUE_ROW
        End If
        
        If main_param.IsSkipBlank() = True And copy_data.GetValue() = "" Then
            '転記元が空の場合はスキップするフラグが真 かつ 転記元が空なので転記しない
            'Common.WriteLog "Copy Value is empty!" & vbCrLf & _
            '                "row=" & row & vbCrLf & _
            '                "keyword=" & keyword
            GoTo CONTINUE_ROW
        End If
        
        Dim FIND_ROW As Long: FIND_ROW = 1
        
        Do
            '指定列の全行を指定ワードで検索し、ヒットした行番号を取得する
            found_row = Common.FindRowByKeywordFromWorksheet( _
                           ws, _
                           sub_param.GetDstFindClm(), _
                           FIND_ROW, _
                           keyword _
                        )
        
            If found_row = 0 Then
                '見つからない!
                'Common.WriteLog "Search keyword is not found!" & vbCrLf & _
                '                "row=" & row & vbCrLf & _
                '                "keyword=" & keyword
                '無視
                Exit Do
            End If
            
            '見つかったので転記
            Set trans_rng = ws.Range(sub_param.GetDstTranClm() & found_row)
            
            Dim src_val As String: src_val = copy_data.GetValue()
            trans_rng.value = src_val
            
            If last_row = found_row Then
                '最終行なのでループを抜ける
                Exit Do
            End If
            
            '見つかった行の次行を再検索
            FIND_ROW = found_row + 1
                  
        Loop
        
CONTINUE_ROW:
        
    Next row
    
    If main_param.IsNotClose() = False Then
        'DSTファイルを保存して閉じる
        Common.SaveAndCloseBook (book_name)
    End If
    
    Common.WriteLog "Transcription E"
End Sub

'転記する(複数行コピー)
Private Sub TranscriptionForMultiCopy(ByRef sub_param As subparam, ByRef copy_datas() As CopyData)
    Common.WriteLog "TranscriptionForMultiCopy S"
    
    Dim ws As Worksheet
    Dim book_name As String
    
    Dim i As Long
    Dim j As Long
    
    Dim mgr As MultiCopyDataManager
    Dim keyword_list() As String
    Dim keyword As String
    Dim find_start_row As Long
    Dim find_end_row As Long
    Dim found_row As Long
    Dim keyword_rows_cnt As Long
    Dim value_list() As String
    Dim find_last_row_num As Long
    
    'DSTファイルパスのDSTシート名を開く
    Set ws = Common.GetSheet(sub_param.GetDstFilePath(), sub_param.GetDstSheetName(), False, True)
    book_name = Common.GetFileName(sub_param.GetDstFilePath())
    
    'キーワード検索範囲の最終行を取得する
    find_end_row = Common.GetLastRowFromWorksheet(ws, sub_param.GetDstFindClm())
    
    '複数行マネージャを生成
    Set mgr = New MultiCopyDataManager
    mgr.Init sub_param, copy_datas
    
    'キーワードリストを取得
    keyword_list = mgr.GetKeywordList()
    
    'コピー元キーワード数分ループ
    For i = 0 To UBound(keyword_list)
        keyword = keyword_list(i)
        
        If keyword = "" Then
            'コピー元キーワードが空なので無視
            GoTo CONTINUE_I
        End If
        
        find_start_row = 1

FIND_ROW:
        'キーワードを検索し､ヒットした行番号を取得する
        found_row = Common.FindRowByKeywordFromWorksheet( _
                       ws, sub_param.GetDstFindClm(), _
                       find_start_row, keyword, find_end_row _
                    )
    
        If found_row = 0 Then
            '見つからないので無視
            GoTo CONTINUE_I
        End If

        'キーワードが見つかった
        
        keyword_rows_cnt = mgr.GetKeywordCount(keyword)
        value_list = mgr.GetValues(keyword)
        
        If keyword_rows_cnt = 1 Then
            'コピー元キーワードが1つの場合
        
            '転記する(1行のみ)
            UpdateCellValue ws, sub_param.GetDstTranClm(), found_row, value_list(0)
            
            find_start_row = found_row + 1
        Else
            'コピー元キーワードが複数行の場合
            
            If IsInsertedRowBeforeSubParam(sub_param) = False Then
                'コピー元行数分、現在行の下に行を挿入する
                Common.InsertRows ws, found_row, keyword_rows_cnt - 1
                
                '挿入済のフラグを立てる
                mgr.SetIsInserted (True)
            End If
            
            '転記する(複数行)
            UpdateMultiCellValues ws, sub_param, keyword, found_row, value_list
            
            '挿入した行数分、検索範囲を更新する
            find_start_row = found_row + keyword_rows_cnt
            find_end_row = find_end_row + keyword_rows_cnt - 1

        End If
        
        find_last_row_num = Common.GetLastRowFromWorksheet(ws, sub_param.GetDstFindClm())
        
        If find_last_row_num < find_start_row Then
            '最終行に達した
            GoTo CONTINUE_I
        End If
        
        GoTo FIND_ROW
        
CONTINUE_I:
        
    Next i
    
    If main_param.IsNotClose() = False Then
        'DSTファイルを保存して閉じる
        Common.SaveAndCloseBook (book_name)
    End If
    
    Common.WriteLog "TranscriptionForMultiCopy E"
End Sub

'転記先のセルを更新する
Private Sub UpdateCellValue(ByRef ws As Worksheet, ByVal clm_name As String, ByVal found_row As Long, ByRef value As String)
    Common.WriteLog "UpdateCellValue S"

    Dim rng As Range
    
    If main_param.IsSkipBlank() = True And value = "" Then
        '転記元が空の場合はスキップするフラグが真 かつ 転記元が空なので転記しない
        Common.WriteLog "UpdateCellValue E-1"
        Exit Sub
    End If
    
    Set rng = ws.Range(clm_name & found_row)
    rng.value = value
    
    Set rng = Nothing
    Common.WriteLog "UpdateCellValue E"
End Sub

'転記先のセルを更新する(複数行)
Private Sub UpdateMultiCellValues(ByRef ws As Worksheet, ByRef sub_param As subparam, ByVal keyword As String, ByVal found_row As Long, ByRef value_list() As String)
    Common.WriteLog "UpdateMultiCellValues S"

    If Common.IsEmptyArray(value_list) = True Then
        Common.WriteLog "UpdateMultiCellValues E-1"
        Exit Sub
    End If
    
    Dim i As Long
    For i = 0 To UBound(value_list)
        '転記する(1行のみ)
        
        'キーワード
        UpdateCellValue ws, sub_param.GetDstFindClm(), found_row + i, keyword
        
        '転記元の値
        UpdateCellValue ws, sub_param.GetDstTranClm(), found_row + i, value_list(i)
    Next i
    
    Common.WriteLog "UpdateMultiCellValues E"
End Sub

'同じシートを対象としたSubParamにおいて既に複数行転記で行を挿入済か?
Private Function IsInsertedRowBeforeSubParam(ByRef sub_param As subparam) As Boolean
    Common.WriteLog "IsInsertedRowBeforeSubParam S"

    Dim i As Long
    Dim before_sub_param As subparam
    Dim ret As Boolean: ret = False

    For i = 0 To UBound(sub_params)
        Set before_sub_param = sub_params(i)
        
        If before_sub_param.GetDstFilePath() = sub_param.GetDstFilePath() And _
           before_sub_param.GetDstSheetName() = sub_param.GetDstSheetName() And _
           before_sub_param.GetDstFindClm() = sub_param.GetDstFindClm() And _
           before_sub_param.GetDstTranClm() <> sub_param.GetDstTranClm() And _
           before_sub_param.IsDstMultiCopy() = True And _
           before_sub_param.IsDstRowInserted() = True Then
           ret = True
           Exit For
        End If

    Next i

    IsInsertedRowBeforeSubParam = ret
    Common.WriteLog "IsInsertedRowBeforeSubParam E"
End Function

