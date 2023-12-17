Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String
Private DQ As String

'パラメータ
Private main_param As MainParam

'--------------------------------------------------------
'メイン処理
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    'パラメータのチェックと収集を行う
    CheckAndCollectParam
    
    '解析して出力する
    Dim ret() As String
    ret = Parse(main_param.GetSrcPath())
    
    If Common.IsExistsFile(main_param.GetDstPath()) = False Then
        Common.CreateFile main_param.GetDstPath()
    End If
    
    OutTsvFile ret, main_param.GetDstPath()

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
    
    Common.WriteLog "CheckAndCollectParam E"
End Sub

Private Function Parse(ByVal src_path As String) As String()
    Common.WriteLog "Parse S"
     
    Dim raw_contents As String
    Dim contents() As String
    Dim ret() As String
    Dim i As Long
    Dim ret_row As Long: ret_row = 0
    Dim line As String
    Dim out_row_cnt As Long
    Dim table_name As String
    Dim clm_name As String
    Dim clm_type As String
    Dim clm_info As String
    Dim clm_constraint As String
    
    If Common.IsSJIS(src_path) = True Then
        raw_contents = Common.ReadTextFileBySJIS(src_path)
    Else
        raw_contents = Common.ReadTextFileByUTF8(src_path)
    End If
    
    contents = Split(raw_contents, vbCrLf)
    contents = Common.DeleteEmptyArray(contents)
    
    out_row_cnt = 0
    table_name = ""
    
    For i = LBound(contents) To UBound(contents)
    
        line = contents(i)
        
        If table_name <> "" Then
            GoTo FOUND_TABLE_NAME
        End If
        
        ' まずはCREATE TABLEを探す
        If Common.IsMatchByRegExp(line, "^CREATE TABLE ", True) = False Then
            GoTo CONTINUE_ROW
        End If
        
        ' 発見したのでテーブル名を取得する
        table_name = GetTableName(line)
        
        GoTo CONTINUE_ROW
        
FOUND_TABLE_NAME:
        ' カラム名等を取得する
        clm_name = GetColumnName(line)
        
        If clm_name = "" Then
            table_name = ""
            GoTo CONTINUE_ROW
        End If
        
        clm_type = GetColumnType(line)
        clm_info = GetColumnInfo(line)
        clm_constraint = GetColumnConstraint(line)
        
        ReDim Preserve ret(ret_row)
        ret(ret_row) = table_name & vbTab & clm_name & vbTab & clm_type & vbTab & clm_info & vbTab & clm_constraint
        
        ret_row = ret_row + 1
        
CONTINUE_ROW:
    
    Next i
    
    Parse = ret

    Common.WriteLog "Parse E"
End Function

Private Function GetTableName(ByVal line As String) As String
    Common.WriteLog "GetTableName S"
    
    Dim ret() As String
    ret = Common.GetMatchByRegExp(line, "\[.*\]", False)
    
    ret = Common.DeleteEmptyArray(ret)
    
    If Common.IsEmptyArray(ret) Then
        GetTableName = ""
        Common.WriteLog "GetTableName E-1"
        Exit Function
    End If
    
    GetTableName = Replace(Replace(ret(0), "[", ""), "]", "")
    
    Common.WriteLog "GetTableName E"
End Function

Private Function GetColumnName(ByVal line As String) As String
    Common.WriteLog "GetColumnName S"
    
    Dim ret() As String
    ret = Common.GetMatchByRegExp(line, "\[(.*?)\]", False)
    
    ret = Common.DeleteEmptyArray(ret)
    
    If Common.IsEmptyArray(ret) Then
        GetColumnName = ""
        Common.WriteLog "GetColumnName E-1"
        Exit Function
    End If
    
    GetColumnName = Replace(Replace(ret(0), "[", ""), "]", "")
    
    Common.WriteLog "GetColumnName E"
End Function

Private Function GetColumnType(ByVal line As String) As String
    Common.WriteLog "GetColumnType S"
    
    Dim ret() As String
    ret = Common.GetMatchByRegExp(line, "\[(.*?)\]", False)
    
    ret = Common.DeleteEmptyArray(ret)
    
    If Common.IsEmptyArray(ret) Then
        GetColumnType = ""
        Common.WriteLog "GetColumnType E-1"
        Exit Function
    End If
    
    GetColumnType = Replace(Replace(ret(1), "[", ""), "]", "")
    
    Common.WriteLog "GetColumnType E"
End Function

Private Function GetColumnInfo(ByVal line As String) As String
    Common.WriteLog "GetColumnInfo S"
    
    Dim ret() As String
    ret = Common.GetMatchByRegExp(line, "\(.*\)", False)
    
    ret = Common.DeleteEmptyArray(ret)
    
    If Common.IsEmptyArray(ret) Then
        GetColumnInfo = ""
        Common.WriteLog "GetColumnInfo E-1"
        Exit Function
    End If
    
    GetColumnInfo = Replace(Replace(ret(0), "(", ""), ")", "")
    
    Common.WriteLog "GetColumnInfo E"
End Function

Private Function GetColumnConstraint(ByVal line As String) As String
    Common.WriteLog "GetColumnConstraint S"
    
    Dim ret As String: ret = Common.GetStringLastChar(line, ")")
    If ret = "" Then
        ret = Common.GetStringLastChar(line, "]")
    End If
    
    GetColumnConstraint = Trim(ret)
    
    Common.WriteLog "GetColumnConstraint E"
End Function

Private Sub OutTsvFile(ByRef datas() As String, ByVal path As String)
    Common.WriteLog "OutTsvFile S"
    
    Common.SaveToFileFromStringArray path, datas
    
    Common.WriteLog "OutTsvFile E"
End Sub
