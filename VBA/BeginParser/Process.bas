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

Private parse_datas() As String

'--------------------------------------------------------
'メイン処理
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    Erase targets
    Erase results
    Erase parse_datas
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
            
            'ファイルパスで始まり、" Begin "を含む行だけ収集する
            If Common.IsMatchByRegExp(line, "^[a-zA-Z]:\\", True) = True And _
               InStr(line, " Begin ") > 0 Then
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
    Dim begin_codes() As String
    
    Set result = New ParseResult
    result.Init target, main_param.GetFormatType()
    
    'ここから先はVBコードのパース処理
    
    'Begin～Endまでのコードを取得する
    GetBeginCodes result

    If result.GetBeginCodesCount() = 0 Then
        'Beginは検出したがEndが検出出来なかった場合(桁位置がズレている可能性大)
        Common.WriteLog "ParseCode E1"
        Exit Sub
    End If

    '1行ずつパースして、結果オブジェクトを作成する
    ParseBeginCode result
    
    parse_datas = Common.DeleteEmptyArray(parse_datas)
    result.SetBeginMembers parse_datas
    Erase parse_datas
    
    If result.GetBeginMembersCount() = 0 Then
        'Begin～Endは検出したがメソッド・プロパティが未使用の場合
        Common.WriteLog "ParseCode E2"
        Exit Sub
    End If
    
    ReDim Preserve results(result_cnt)
    Set results(result_cnt) = result
    result_cnt = result_cnt + 1
    
    Common.WriteLog "ParseCode E"
End Sub

'Begin～Endまでのコードを配列で返す
'→入れ子のBeginは無視する
Private Sub GetBeginCodes( _
    ByRef result As ParseResult _
)
    Common.WriteLog "GetBeginCodes S"
    
    Dim raw_contents() As String
    Dim begin_codes() As String
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
    
    'Begin～Endまでの行を配列に入れる
    For i = result.GetRowNum() - 1 To UBound(raw_contents)
        line = raw_contents(i)

        If Common.IsCommentCode(line, ext) = True Then
            'コメント行なので次の行へ
            GoTo CONTINUE
        End If

        '右コメントを除去しておく
        line = Common.RemoveRightComment(line, ext)
        
        If Common.IsMatchByRegExp(line, "^Begin .*$", True) = True Then
        
            'Beginを検出
            
            clm_wk = Common.FindFirstCasePosition(line)
            
            If first_clm = 0 Then
                '最初に検出したBeginの桁位置を保持しておく
                first_clm = clm_wk
            End If
        
        ElseIf Common.IsMatchByRegExp(line, "^End$", True) = True Then
        
            'End Beginを検出
            
            clm_wk = Common.FindFirstCasePosition(line)
            If clm_wk = first_clm Then
                'Grep結果のBeginに対応するEndを発見したので終了
                ReDim Preserve begin_codes(cnt)
                begin_codes(cnt) = line
                is_find = True
                Exit For
            End If
                    
        End If
    
        ReDim Preserve begin_codes(cnt)
        begin_codes(cnt) = line
        cnt = cnt + 1

CONTINUE:
    
    Next i
    
    If is_find = False Then
        Dim err_msg As String
    
        err_msg = "Grep結果のBeginに対応するEndが見つかりません (target=" & result.GetTarget() & ")"
    
        If Common.ShowYesNoMessageBox( _
            "[GetBeginCodes]でエラーが発生しました。処理を続行しますか?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[GetBeginCodes] エラー! (err_msg=" & err_msg & ")"
        End If
        
        Common.WriteLog err_msg
        Common.WriteLog "GetBeginCodes E1"
        Exit Sub
    End If
    
    result.SetBeginCodes begin_codes
    
    Common.WriteLog "GetBeginCodes E"
End Sub

'1行ずつパースして、結果オブジェクトを作成する
Private Sub ParseBeginCode( _
    ByRef result As ParseResult _
)
    Common.WriteLog "ParseBeginCode S"

    Parse result, result.GetBeginCodes()(0), result.GetBeginClass()

    Common.WriteLog "ParseBeginCode E"
End Sub

'階層探索
Private Sub Parse( _
    ByRef result As ParseResult, _
    ByVal name As String, _
    ByVal path As String _
)
    Common.WriteLog "Parse S"

    AppendPropertyPath result, name, path
    
    Dim sub_begins() As String
    sub_begins = GetSubBeginList(result, name)
    
    If Common.IsEmptyArray(sub_begins) = True Then
        Common.WriteLog "Parse E1"
        Exit Sub
    End If
    
    Dim i As Long
    Dim sub_name As String
    Dim sub_path As String
    
    For i = 0 To UBound(sub_begins)
        sub_name = sub_begins(i)
        sub_path = path & "/" & Replace(Replace(Trim(sub_name), "BeginProperty", ""), "Begin", "")
        
        '再帰
        Parse result, sub_name, sub_path
        
    Next i

    Common.WriteLog "Parse E"
End Sub

'現在の階層にあるプロパティ(Key = Value)を列挙して、結合パスの末尾にデリミタ付きで結合する
Private Sub AppendPropertyPath( _
    ByRef result As ParseResult, _
    ByVal name As String, _
    ByVal path As String _
)
    Common.WriteLog "AppendPropertyPath S"

    Const REG = "^(?!.*(Begin |BeginProperty |EndProperty$|End$)).*$"
    Dim contents() As String
    contents = GetSubBeginContents(result, name)
    
    Dim i As Long
    Dim line As String
    Dim member() As String
    Dim cnt As Long: cnt = 0
    
    'プロパティだけ収集
    For i = 0 To UBound(contents)
        line = contents(i)
        
        If Common.IsMatchByRegExp(line, REG, True) = True Then
            ReDim Preserve member(cnt)
            member(cnt) = line
            cnt = cnt + 1
        End If
    Next i

    '最小桁位置を取得
    Dim min_clm As Long: min_clm = GetMinColumn(member)
    
    '現在の階層のプロパティだけにする
    Dim member_current() As String
    Dim clm_wk As Long
    cnt = 0
    
    For i = 0 To UBound(member)
        line = member(i)
        
        clm_wk = Common.FindFirstCasePosition(line)
        
        If clm_wk <= min_clm Then
            ReDim Preserve member_current(cnt)
            member_current(cnt) = path & "/" & Trim(line)
            cnt = cnt + 1
        End If
    Next i
    
    '最後にマージ
    parse_datas = Common.MergeArray(parse_datas, member_current)

    Common.WriteLog "AppendPropertyPath E"
End Sub
    
'最小桁位置を返す
'TODO: Common化
Private Function GetMinColumn(ByRef ary() As String) As Long
    Common.WriteLog "GetMinColumn S"

    Dim i As Long
    Dim clm_wk As Long
    Dim line As String
    Dim min_clm As Long: min_clm = -1
    
    For i = 0 To UBound(ary)
        line = ary(i)
        clm_wk = Common.FindFirstCasePosition(line)
        
        If min_clm = -1 Then
            min_clm = clm_wk
        Else
            If min_clm > clm_wk Then
                '最小を発見
                min_clm = clm_wk
            End If
        End If
        
        
    Next i
    
    GetMinColumn = min_clm

    Common.WriteLog "GetMinColumn E"
End Function

'現在の階層にあるサブ階層("Begin" or "BeginPrpperty"で始まる)を列挙する
Private Function GetSubBeginList( _
    ByRef result As ParseResult, _
    ByVal name As String _
) As String()
    Common.WriteLog "GetSubBeginList S"

    Const REG = "Begin |BeginProperty "
    
    Dim contents() As String
    contents = GetSubBeginContents(result, name)
    
    Dim i As Long
    Dim line As String
    Dim member() As String
    Dim cnt As Long: cnt = 0
    
    Common.WriteLog "contents=" & CStr(UBound(contents))
    
    'Begin, BeginPropertyだけを収集
    For i = 0 To UBound(contents)
        If i = 0 Then
            '1行目は無視
            GoTo CONTINUE
        End If
        
        line = contents(i)
        
        If Common.IsMatchByRegExp(line, REG, True) = True Then
            ReDim Preserve member(cnt)
            member(cnt) = line
            cnt = cnt + 1
        End If
CONTINUE:
    Next i
    
    If Common.IsEmptyArray(member) = True Then
        GetSubBeginList = member
        Common.WriteLog "GetSubBeginList E1"
        Exit Function
    End If
    
    '最小桁数を取得
    Dim min_clm As Long: min_clm = GetMinColumn(member)
    
    '現在の階層のBegin, BeginPropertyだけにする
    Dim member_current() As String
    Dim clm_wk As Long
    cnt = 0
    
    For i = 0 To UBound(member)
        line = member(i)
        
        clm_wk = Common.FindFirstCasePosition(line)
        
        If clm_wk <= min_clm Then
            ReDim Preserve member_current(cnt)
            member_current(cnt) = line
            cnt = cnt + 1
        End If
    Next i
    
    GetSubBeginList = member_current

    Common.WriteLog "GetSubBeginList E"
End Function

'指定された階層の内容を返す
Private Function GetSubBeginContents( _
    ByRef result As ParseResult, _
    ByVal name As String _
) As String()
    Common.WriteLog "GetSubBeginContents S"

    Dim i As Long
    Dim line As String
    Dim ext As String: ext = result.GetExtension()
    Dim first_clm As Long: first_clm = -1
    Dim clm_wk As Long
    Dim is_find As Boolean: is_find = False
    Dim contents() As String
    Dim cnt As Long: cnt = 0
    Dim end_word As String
    
    Dim begin_type As Integer   '0=Begin, 1=BeginProperty
    
    If Left(Trim(name), 6) = "Begin " Then
        begin_type = 0
        end_word = "^End$"
    ElseIf Left(Trim(name), 14) = "BeginProperty " Then
        begin_type = 1
        end_word = "^EndProperty$"
    Else
        Err.Raise 53, , "キーワードが見つかりません! (target=" & result.GetTarget() & ")"
    End If
    
    'Begin～Endまでの行を配列に入れる
    For i = 0 To UBound(result.GetBeginCodes())
        line = result.GetBeginCodes()(i)
        
        If Common.IsCommentCode(line, ext) = True Then
            'コメント行なので次の行へ
            GoTo CONTINUE
        End If
    
        '右コメントを除去しておく
        line = Common.RemoveRightComment(line, ext)
        
        If first_clm = -1 And line = name Then
            first_clm = Common.FindFirstCasePosition(line)
        End If
    
        If first_clm = -1 Then
            '対象を見つけていないので無視
            GoTo CONTINUE
        End If
    
        If Common.IsMatchByRegExp(Trim(line), end_word, True) = True Then
            clm_wk = Common.FindFirstCasePosition(line)
            If clm_wk = first_clm Then
                'Beginに対するEndを発見したので終了
                ReDim Preserve contents(cnt)
                contents(cnt) = line
                is_find = True
                Exit For
            End If
        End If
    
        ReDim Preserve contents(cnt)
        contents(cnt) = line
        cnt = cnt + 1
    
CONTINUE:
    
    Next i
    
    If is_find = False Then
        Dim err_msg As String
        
        err_msg = "Grep結果のBeginに対応するEndが見つかりません (target=" & result.GetTarget() & ")"
        
        If Common.ShowYesNoMessageBox( _
            "[GetSubBeginContents]でエラーが発生しました。処理を続行しますか?" & vbCrLf & _
            "err_msg=" & err_msg _
            ) = False Then
            Err.Raise 53, , "[GetSubBeginContents] エラー! (err_msg=" & err_msg & ")"
        End If
        
        Common.WriteLog err_msg
        Common.WriteLog "GetSubBeginContents E1"
        Exit Function
    End If
    
    GetSubBeginContents = contents

    Common.WriteLog "GetSubBeginContents E"
End Function

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
    ws.Range("A1").value = "Begin句の解析結果"
    
    '列名を追加
    ws.Range("A3").value = "GREP結果"
    ws.Range("B3").value = "ファイルパス"
    ws.Range("C3").value = "プロパティ"
    ws.Range("D3").value = "値"
    ws.Range("E3").value = "ルート"
    ws.Range("F3").value = "階層1"
    ws.Range("G3").value = "階層2"
    ws.Range("H3").value = "階層3"
    ws.Range("I3").value = "階層4"
    ws.Range("J3").value = "階層5"
    ws.Range("K3").value = "階層6"
    ws.Range("L3").value = "階層7"
    ws.Range("M3").value = "階層8"
    ws.Range("N3").value = "階層9"
    ws.Range("O3").value = "階層10"
    
    
    Const START_ROW = 4
    Dim i As Long
    Dim j As Long
    Dim offset_row As Long: offset_row = 0
    Dim result As ParseResult
    Dim row As Long: row = START_ROW
    Dim members() As String
    
    Dim cnt As Long
    
    Dim key_ As String
    Dim val_ As String
    Dim k1 As String
    Dim k2 As String
    Dim k3 As String
    Dim k4 As String
    Dim k5 As String
    Dim k6 As String
    Dim k7 As String
    Dim k8 As String
    Dim k9 As String
    
    Dim key_val() As String
    
    
    '結果オブジェクトリストでループ
    For i = 0 To UBound(results)
        Set result = results(i)
        
        If result.GetBeginMembersCount() = 0 Then
            GoTo CONTINUE
        End If
        
        '結果オブジェクトの内容を記載
        members = result.GetBeginMembers()
        
        For j = 0 To UBound(members)
            Dim items() As String: items = Split(members(j), "/")
            
            If Common.IsEmptyArray(items) = True Then
                GoTo CONTINUE_J
            End If
            
            cnt = UBound(items)
            
            key_ = ""
            val_ = ""
            k1 = ""
            k2 = ""
            k3 = ""
            k4 = ""
            k5 = ""
            k6 = ""
            k7 = ""
            k8 = ""
            k9 = ""
        
            ws.Cells(row + i + offset_row + j, 1).value = result.GetTarget()
            ws.Cells(row + i + offset_row + j, 2).value = result.GetFilePath()
            
            key_val = Split(items(cnt), "=")
            
            If Common.IsEmptyArray(key_val) = True Or UBound(key_val) = 0 Then
                GoTo CONTINUE_J
            End If
            
            key_ = Trim(key_val(0))
            val_ = Trim(key_val(1))
            
            If cnt = 1 Then
            
            ElseIf cnt = 2 Then
                k1 = items(1)
            ElseIf cnt = 3 Then
                k1 = items(1)
                k2 = items(2)
            ElseIf cnt = 4 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
            ElseIf cnt = 5 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
                k4 = items(4)
            ElseIf cnt = 6 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
                k4 = items(4)
                k5 = items(5)
            ElseIf cnt = 7 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
                k4 = items(4)
                k5 = items(5)
                k6 = items(6)
            ElseIf cnt = 8 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
                k4 = items(4)
                k5 = items(5)
                k6 = items(6)
                k7 = items(7)
            ElseIf cnt = 9 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
                k4 = items(4)
                k5 = items(5)
                k6 = items(6)
                k7 = items(7)
                k8 = items(8)
            ElseIf cnt = 10 Then
                k1 = items(1)
                k2 = items(2)
                k3 = items(3)
                k4 = items(4)
                k5 = items(5)
                k6 = items(6)
                k7 = items(7)
                k8 = items(8)
                k9 = items(9)
            End If
            
            ws.Cells(row + i + offset_row + j, 3).value = key_
            ws.Cells(row + i + offset_row + j, 4).value = val_
            ws.Cells(row + i + offset_row + j, 5).value = result.GetBeginClass()
            
            ws.Cells(row + i + offset_row + j, 6).value = k1
            ws.Cells(row + i + offset_row + j, 7).value = k2
            ws.Cells(row + i + offset_row + j, 8).value = k3
            ws.Cells(row + i + offset_row + j, 9).value = k4
            ws.Cells(row + i + offset_row + j, 10).value = k5
            ws.Cells(row + i + offset_row + j, 11).value = k6
            ws.Cells(row + i + offset_row + j, 12).value = k7
            ws.Cells(row + i + offset_row + j, 13).value = k8
            ws.Cells(row + i + offset_row + j, 14).value = k9
        
CONTINUE_J:
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

