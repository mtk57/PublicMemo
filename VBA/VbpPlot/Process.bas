Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String
Private DQ As String

'パラメータ
Public main_param As MainParam

Private vbprj_files() As String

'--------------------------------------------------------
'メイン処理
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    Dim err_msg As String
    Dim ws As Worksheet
    Dim row As Long, col As Long
    Dim clm_dict As Object
    
    
    SEP = Application.PathSeparator
    DQ = Chr(34)
    
    'パラメータのチェックと収集を行う
    CheckAndCollectParam
    
    ' 新しいシートを作成
    Set ws = ThisWorkbook.Sheets.Add
    
    ' 結果を格納する辞書を作成
    Set clm_dict = CreateObject("Scripting.Dictionary")
    
    ' VBPファイルを検索して処理
    SearchVBPFiles main_param.GetSrcDirPath(), ws, clm_dict
    
    ' 結果を出力
    OutputResults ws, clm_dict

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

Private Sub SearchVBPFiles(ByVal folderPath As String, ByRef ws As Worksheet, ByRef clmDict As Object)
    Common.WriteLog "SearchVBPFiles S"
    
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim subFolder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' 現在のフォルダ内のファイルを処理
    For Each file In folder.files
        If LCase(fso.GetExtensionName(file.Name)) = "vbp" Then
            ProcessVBPFile file.path, ws, clmDict
        End If
    Next file
    
    ' サブフォルダを再帰的に処理
    For Each subFolder In folder.SubFolders
        SearchVBPFiles subFolder.path, ws, clmDict
    Next subFolder

    Common.WriteLog "SearchVBPFiles E"
End Sub

Private Sub ProcessVBPFile(ByVal filePath As String, ByRef ws As Worksheet, ByRef clmDict As Object)
    Common.WriteLog "ProcessVBPFile S"
    
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts() As String
    Dim key As String, value As String
    Dim row As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1) ' 1 = ForReading
    
    ' ファイルパスを出力
    row = ws.Cells(ws.Rows.count, "A").End(xlUp).row + 1
    ws.Cells(row, 1).value = filePath
    
    ' ファイルの内容を処理
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        parts = Split(line, "=", 2)
        If UBound(parts) = 1 Then
            key = Trim(parts(0))
            
            If main_param.IsExistIgnoreKey(key) = False Then
                GoTo CONTINUE
            End If
            
            value = Trim(parts(1))
            
            ' キーが存在しない場合、新しい列を追加
            If Not clmDict.Exists(key) Then
                clmDict.Add key, clmDict.count + 1
            End If
            
            ' 値を適切なセルに設定
            ws.Cells(row, clmDict(key) + 1).value = value
        End If
CONTINUE:
    Loop
    
    ts.Close

    Common.WriteLog "ProcessVBPFile E"
End Sub

Private Sub OutputResults(ByRef ws As Worksheet, ByRef clmDict As Object)
    Common.WriteLog "OutputResults S"
    
    Dim key As Variant
    Dim col As Long
    
    ' ヘッダーを出力
    For Each key In clmDict.Keys
        col = clmDict(key)
        ws.Cells(1, col + 1).value = key
    Next key
    
    '' セルの書式を設定
    'With ws.UsedRange
    '    .EntireColumn.AutoFit
    '    .HorizontalAlignment = xlLeft
    '    .VerticalAlignment = xlTop
    '    .WrapText = True
    'End With
    
    Common.WriteLog "OutputResults E"
End Sub
