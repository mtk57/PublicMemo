Attribute VB_Name = "Main"
Option Explicit

' フォルダ配下の全ファイル情報をExcelシートに取得するマクロ
Sub GetAllFileInfo()
    Dim folderPath As String
    Dim outputSheet As Worksheet
    Dim fso As Object
    Dim folder As Object
    Dim row As Long
    
    ' FileSystemObjectを作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' フォルダ選択ダイアログを表示
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "ファイル情報を取得するフォルダを選択してください"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "フォルダが選択されませんでした。", vbExclamation
            Exit Sub
        End If
    End With
    
    ' フォルダが存在するか確認
    If Not fso.FolderExists(folderPath) Then
        MsgBox "指定されたフォルダが見つかりません: " & folderPath, vbExclamation
        Exit Sub
    End If
    
    'シートを追加
    Dim sheet_name_ As String: sheet_name_ = Common.GetNowTimeString()
    Common.AddSheet ActiveWorkbook, sheet_name_
    
    ' 出力先シートの準備
    Application.ScreenUpdating = False
    
    ' 既存のシートがあれば削除し、新しいシートを作成
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set outputSheet = ThisWorkbook.Sheets(sheet_name_)
    
    ' ヘッダーの設定
    outputSheet.Cells(1, 1).value = "ファイルパス"
    outputSheet.Cells(1, 2).value = "更新年月日"
    outputSheet.Cells(1, 3).value = "更新時分秒"
    outputSheet.Cells(1, 4).value = "ファイルサイズ"
    
    ' ヘッダー行の書式設定は省略

    
    ' 初期行番号
    row = 2
    
    ' フォルダを取得
    Set folder = fso.GetFolder(folderPath)
    
    ' フォルダ内のファイル情報を収集
    ProcessFiles folder, outputSheet, row
    
    ' 列幅の自動調整のみ実施
    outputSheet.Columns("A:D").AutoFit
    
    Application.ScreenUpdating = True
    
    MsgBox "ファイル情報の取得が完了しました。" & vbCrLf & _
           "ファイル数: " & (row - 2), vbInformation
End Sub

' フォルダ内のファイルを再帰的に処理する関数
Private Sub ProcessFiles(folder As Object, outputSheet As Worksheet, ByRef row As Long)
    Dim file As Object
    Dim subfolder As Object
    Dim lastModified As Date
    
    ' フォルダ内のすべてのファイルを処理
    For Each file In folder.Files
        ' ファイル情報を取得
        lastModified = file.DateLastModified
        
        ' シートに出力
        outputSheet.Cells(row, 1).value = file.path
        outputSheet.Cells(row, 2).value = Format(lastModified, "yyyy/mm/dd")
        outputSheet.Cells(row, 3).value = Format(lastModified, "hh:mm:ss")
        outputSheet.Cells(row, 4).value = file.size
        
        ' 行番号を増やす
        row = row + 1
    Next file
    
    ' サブフォルダを再帰的に処理
    For Each subfolder In folder.SubFolders
        ProcessFiles subfolder, outputSheet, row
    Next subfolder
End Sub




