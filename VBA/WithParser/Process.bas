Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String
Private DQ As String

'パラメータ
Private main_param As MainParam
Private sub_param As SubParam

Private targets() As String
Private results() As ParseResult
Private result_cnt As Long

'--------------------------------------------------------
'メイン処理
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    Erase targets
    Erase results
    result_cnt = 0

    'パラメータのチェックと収集を行う
    CheckAndCollectParam
    
    'メインループ
    Dim i As Long
    For i = LBound(targets) To UBound(targets)
        Dim target As String: target = targets(i)
        Common.WriteLog "i=" & i & ":[" & target & "]"
    
        'コードを解析する
        ParseCode target
    Next i
    
    'シートに結果を出力する
    OutputSheet

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
    
    If main_param.GetFormatType() = "sakura" Then
        For i = LBound(grep_result) To UBound(grep_result)
            line = grep_result(i)
            
            'ファイルパスで始まり、" With "を含む行だけ収集する
            If Common.IsMatchByRegExp(line, "^[a-zA-Z]:\\", True) = True And _
               InStr(line, " With ") > 0 Then
                ReDim Preserve targets(cnt)
                targets(cnt) = line
                
                cnt = cnt + 1
            End If
        Next i
    End If
    
    Common.WriteLog "CheckAndCollectParam E"
End Sub

'コードを解析する
Private Sub ParseCode( _
    ByVal target As String _
)
    Common.WriteLog "ParseCode S"
    
    Dim result As ParseResult
    Dim with_codes() As String
    
    Set result = New ParseResult
    result.Init target, main_param.GetFormatType()
    
    'ここから先はVBコードのパース処理
    
    'With～End Withまでのコードを取得する
    GetWithCodes result

    If result.GetWithCodesCount() = 0 Then
        'Withは検出したがEnd Withが検出出来なかった場合(桁位置がズレている可能性大)
        Common.WriteLog "ParseCode E1"
        Exit Sub
    End If

    '1行ずつパースして、結果オブジェクトを作成する
    ParseWithCode result
    
    If result.GetWithMembersCount() = 0 Then
        'With～End Withは検出したがメソッド・プロパティが未使用の場合
        Common.WriteLog "ParseCode E2"
        Exit Sub
    End If
    
    ReDim Preserve results(result_cnt)
    Set results(result_cnt) = result
    result_cnt = result_cnt + 1
    
    Common.WriteLog "ParseCode E"
End Sub

'With～End Withまでのコードを配列で返す
'→入れ子のWithは無視する
Private Sub GetWithCodes( _
    ByRef result As ParseResult _
)
    Common.WriteLog "GetWithCodes S"
    
    Dim raw_contents() As String
    Dim with_codes() As String
    Dim i As Long
    Dim cnt As Long: cnt = 0
    Dim line As String
    Dim ext As String: ext = result.GetExtension()
    Dim is_find As Boolean: is_find = False
    Dim clm_wk As Long
    Dim first_clm As Long: first_clm = 0
    Dim is_ignore As Boolean: is_ignore = False
    
    'ファイルパスのファイルを開く
    raw_contents = GetTargetContents(result)
    
    'With～End Withまでの行を配列に入れる
    For i = result.GetRowNum() - 1 To UBound(raw_contents)
        line = raw_contents(i)

        If Common.IsCommentCode(line, ext) = True Then
            'コメント行なので次の行へ
            GoTo CONTINUE
        End If

        '右コメントを除去しておく
        line = Common.RemoveRightComment(line, ext)
        
        If Common.IsMatchByRegExp(line, "^ *With .*$", True) = True Then
        
            'Withを検出
            
            clm_wk = Common.FindFirstCasePosition(line)
            
            If first_clm = 0 Then
                '最初に検出したWithの桁位置を保持しておく
                first_clm = clm_wk
            End If
            
            If clm_wk <> first_clm Then
                '入れ子のWithを検出したので無視
                is_ignore = True
                GoTo CONTINUE
            End If
        
        ElseIf Common.IsMatchByRegExp(line, "^ *End With$", True) = True Then
        
            'End Withを検出
            
            If is_ignore = True Then
                '入れ子のWithの終了を検出
                is_ignore = False
                GoTo CONTINUE
            End If
        
            clm_wk = Common.FindFirstCasePosition(line)
            If clm_wk = first_clm Then
                'Grep結果のWithに対応するEnd Withを発見したので終了
                ReDim Preserve with_codes(cnt)
                with_codes(cnt) = line
                is_find = True
                Exit For
            End If
        
        Else
            
            'With, End With以外の行
            If is_ignore = True Then
                '入れ子のWithの終了を検出していないので無視
                GoTo CONTINUE
            End If
                    
        End If
    
        ReDim Preserve with_codes(cnt)
        with_codes(cnt) = line
        cnt = cnt + 1

CONTINUE:
    
    Next i
    
    If is_find = False Then
        Dim err_msg As String
    
        err_msg = "Grep結果のWithに対応するEnd Withが見つかりません (target=" & result.GetTarget() & ")"
    
        If Common.ShowYesNoMessageBox( _
            "[GetWithCodes]でエラーが発生しました。処理を続行しますか?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[GetWithCodes] エラー! (err_msg=" & err_msg & ")"
        End If
        
        Common.WriteLog err_msg
        Common.WriteLog "GetWithCodes E1"
        Exit Sub
    End If
    
    result.SetWithCodes with_codes
    
    Common.WriteLog "GetWithCodes E"
End Sub

'1行ずつパースして、結果オブジェクトを作成する
Private Sub ParseWithCode( _
    ByRef result As ParseResult _
)
    Common.WriteLog "ParseWithCode S"
    
    Const MEMBER = "(\s|\()\.[a-zA-Z][a-zA-Z0-9_]*"
    
    Dim i As Long
    Dim j As Long
    Dim with_codes() As String
    Dim with_class As String
    Dim temp_ary() As String
    Dim with_members() As String
    Dim line As String
    
    with_codes = result.GetWithCodes()
    
    For i = 0 To UBound(with_codes)
        line = with_codes(i)
        
        Common.WriteLog "[" & i & "]=" & line
        
        If i = 0 Then
            with_class = Trim(Replace(line, "With", ""))
            GoTo CONTINUE
        End If
        
        temp_ary = Common.DeleteEmptyArray(Common.GetMatchByRegExp(line, MEMBER, True))
        If Common.IsEmptyArray(temp_ary) = True Then
            'ドットで始まるメソッド・プロパティが存在しないので次の行へ
            GoTo CONTINUE
        End If
        
        For j = 0 To UBound(temp_ary)
            temp_ary(j) = Replace(Trim(temp_ary(j)), "(", "")
        Next j
        
        with_members = Common.MergeArray(with_members, temp_ary)
    
CONTINUE:
    
    Next i
    
    If Common.IsEmptyArray(with_members) = True Then
        Common.WriteLog "ParseWithCode E1"
        Exit Sub
    End If
    
    result.SetWithClass with_class
    result.SetWithMembers Common.SortAndDistinctArray(Common.DeleteEmptyArray(with_members))

    Common.WriteLog "ParseWithCode E"
End Sub

'シートに結果を出力する
Private Sub OutputSheet()
    Common.WriteLog "OutputSheet S"
    
    If Common.IsEmptyArray(results) = True Then
        Common.WriteLog "OutputSheet E1"
        Exit Sub
    End If
    
    'シートを追加
    Dim sheet_name As String: sheet_name = Common.GetNowTimeString()
    Common.AddSheet ActiveWorkbook, sheet_name
    
    'シートのタイトルを追加
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("A1").value = "With句の解析結果"
    
    '列名を追加
    ws.Range("A3").value = "GREP結果"
    ws.Range("B3").value = "ファイルパス"
    ws.Range("C3").value = "クラス"
    ws.Range("D3").value = "メソッド/プロパティ"
    
    
    Const START_ROW = 4
    Dim i As Long
    Dim j As Long
    Dim offset_row As Long: offset_row = 0
    Dim result As ParseResult
    Dim row As Long: row = START_ROW
    Dim members() As String
    
    '結果オブジェクトリストでループ
    For i = 0 To UBound(results)
        Set result = results(i)
        
        If result.GetWithMembersCount() = 0 Then
            GoTo CONTINUE
        End If
        
        '結果オブジェクトの内容を記載
        members = result.GetWithMembers()
        
        For j = 0 To UBound(members)
            
            If j = 0 Then
                ws.Cells(row + i + offset_row + j, 1).Font.Color = RGB(0, 0, 0)
                ws.Cells(row + i + offset_row + j, 2).Font.Color = RGB(0, 0, 0)
                ws.Cells(row + i + offset_row + j, 3).Font.Color = RGB(0, 0, 0)
                ws.Cells(row + i + offset_row + j, 4).Font.Color = RGB(0, 0, 0)
            Else
                ws.Cells(row + i + offset_row + j, 1).Font.Color = RGB(192, 192, 192)
                ws.Cells(row + i + offset_row + j, 2).Font.Color = RGB(192, 192, 192)
                ws.Cells(row + i + offset_row + j, 3).Font.Color = RGB(192, 192, 192)
                ws.Cells(row + i + offset_row + j, 4).Font.Color = RGB(0, 0, 0)
            End If
        
            ws.Cells(row + i + offset_row + j, 1).value = result.GetTarget()
            ws.Cells(row + i + offset_row + j, 2).value = result.GetFilePath()
            ws.Cells(row + i + offset_row + j, 3).value = result.GetWithClass()
            ws.Cells(row + i + offset_row + j, 4).value = members(j)

        Next j
        
        offset_row = offset_row + UBound(members)
        
CONTINUE:
    Next i
    
    Common.WriteLog "OutputSheet E"
End Sub

'対象ファイルを読み込んで内容を配列で返す
Private Function GetTargetContents( _
    ByRef result As ParseResult _
) As String()
    Common.WriteLog "GetTargetContents S"
    
    Dim raw_contents As String
    Dim contents() As String
    
    'ファイルを開いて、全行を配列に格納する
    If result.GetEncode() = "SJIS" Then
        raw_contents = Common.ReadTextFileBySJIS(result.GetFilePath())
    ElseIf result.GetEncode() = "UTF-8" Then
        raw_contents = Common.ReadTextFileByUTF8(result.GetFilePath())
    Else
        Dim err_msg As String: err_msg = "未サポートのエンコードです" & vbCrLf & _
                  "path=" & result.GetFilePath()
        Common.WriteLog "[GetTargetContents] ★★エラー! err_msg=" & err_msg
        
        If Common.ShowYesNoMessageBox( _
            "[GetTargetContents]でエラーが発生しました。処理を続行しますか?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[GetTargetContents] エラー! (err_msg=" & err_msg & ")"
        End If
        
        Common.WriteLog "GetTargetContents E1"
        GetTargetContents = contents
        Exit Function
    End If
    
    contents = Split(raw_contents, vbCrLf)
    
    GetTargetContents = contents

    Common.WriteLog "GetTargetContents E"
End Function

