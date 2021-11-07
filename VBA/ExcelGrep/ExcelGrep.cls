VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelGrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Base code : https://gist.github.com/TatsuoWatanabe/2ed13d04d9dd686e9b23206c88a7e1a4
Option Explicit

Private objMainSheet_ As Worksheet

Private objTargetPathPos_ As Range
Private objSearchWordPos_ As Range
Private objIgnoreCasePos_ As Range
Private objResultListFirstPos_ As Range

Private shouldSearch_ As Boolean

Private objFSO_ As New FileSystemObject
Private objREG_ As New VBScript_RegExp_55.RegExp
Private objExcelApp_ As Excel.Application

Private Enum ResultColumns
    Path = 1
    Book = 2
    Sheet = 3
    Pos = 4
    Val = 5
End Enum

Private Sub Class_Initialize()
    Set objMainSheet_ = ThisWorkbook.Sheets("ExcelGrep")
    Set objTargetPathPos_ = objMainSheet_.Range("C3")
    Set objSearchWordPos_ = objMainSheet_.Range("C4")
    Set objIgnoreCasePos_ = objMainSheet_.Range("C6")
    Set objResultListFirstPos_ = objMainSheet_.Range("C17")
    shouldSearch_ = False
End Sub

Public Sub ExecSearch()
    Dim objFolder As Folder
    Dim ignoreCase As Boolean
    
    Set objFolder = GetTargetFolder()
    ignoreCase = False
    
    If objIgnoreCasePos_ = "" Or objIgnoreCasePos_ = "区別する（デフォルト）" Then
        ignoreCase = True
    End If
    
    If objFolder Is Nothing Then
        ShowInfoMsgBox ("検索対象フォルダが見つかりませんでした。")
        Exit Sub
    End If
    
    If Trim(objSearchWordPos_.value) = "" Then
        ShowInfoMsgBox ("検索文字列を入力してください。")
        Exit Sub
    End If
    
    If shouldSearch_ = True Then
        ShowInfoMsgBox ("別の検索処理がすでに実行中です。")
        Exit Sub
    End If
    
    objREG_.Global = True
    objREG_.ignoreCase = ignoreCase
    objREG_.Pattern = objSearchWordPos_.value
    
    '検索開始
    shouldSearch_ = True
    Call ClearResultList
    Call SearchFolder(objFolder)
    
    '検索終了
    If Not objExcelApp_ Is Nothing Then
        objExcelApp_.Quit
        Set objExcelApp_ = Nothing
    End If
    
    shouldSearch_ = False
    Call UpdateStatusBar("")
    ShowInfoMsgBox ("検索が終了しました。")
End Sub

Public Sub StopSearch()
    If shouldSearch_ = False Then
        Exit Sub
    End If
    
    If ShowYNMsgBox("検索を中止してもよろしいですか?") <> vbYes Then
        Exit Sub
    End If
    
    shouldSearch_ = False
End Sub

Public Sub ClearResultList()
    'CurrentRegionはアクティブセル領域（※）のこと。
    'Offsetはアクティブセル領域を下に2行ずらす。
    'これで検索結果の領域だけが対象となる。（検索結果の2行上には「表名」と「表の列」の2行があり消してはダメなので）
    '
    '※ アクティブセル領域とは
    '   指定したセルを含む、空白行と空白列で囲まれた、データの入力範囲です。
    '   [CTRL] + [SHIFT] + [*] でも選択できる
    objResultListFirstPos_.CurrentRegion.Offset(2).Delete
End Sub

Private Function GetTargetFolder() As Folder
On Error Resume Next
    Set GetTargetFolder = objFSO_.GetFolder(objTargetPathPos_.value)
End Function

Private Sub SearchFolder(objFolder As Folder)
    Dim objFile As File
    Dim objSubFolder As Folder
    
    If shouldSearch_ = False Then
        Exit Sub
    End If
    
    Call UpdateStatusBar(objFolder.Path)
    
    For Each objFile In objFolder.Files
        Select Case objFSO_.GetExtensionName(objFile.Path)
            Case "xls", "xlsx", "xlsm"
                Call SearchBook(objFile)
        End Select
    Next
    
    For Each objSubFolder In objFolder.SubFolders
        Call SearchFolder(objSubFolder)
    Next
End Sub

Private Sub SearchBook(objFile As File)
    Dim objBook As Workbook
    Dim objSheet As Worksheet
    
    If shouldSearch_ = False Then
        Exit Sub
    End If

On Error Resume Next
    If objExcelApp_ Is Nothing Then
        Set objExcelApp_ = New Excel.Application
        objExcelApp_.Visible = False
        objExcelApp_.ScreenUpdating = False
    End If

    Set objBook = objExcelApp_.Workbooks.Open(Filename:=objFile.Path, ReadOnly:=True)
    If objBook Is Nothing Then
        ShowInfoMsgBox (objFile.Path & vbCrLf & " が開けませんでした。")
        Exit Sub
    End If

On Error GoTo 0
    For Each objSheet In objBook.Worksheets
        Call SearchSheet(objSheet)
    Next
    
    Call objBook.Close(SaveChanges:=False)
End Sub

Private Sub SearchSheet(Sheet As Worksheet)
    Dim objTargetRange As Range
    Dim objCell As Range
    Dim objShape As Shape

    If shouldSearch_ = False Then
        Exit Sub
    End If
    
    'Search Cells
    Set objTargetRange = Sheet.UsedRange.Cells
    For Each objCell In objTargetRange
        DoEvents
        Call UpdateStatusBar(Sheet.Parent.FullName)
        
        If objCell.value <> "" Then
            If objREG_.Test(objCell.value) Then
                '検索文字列を発見
                Call ProcessCell(objCell)
            End If
        End If
    Next
    
    'Search Shapes
    For Each objShape In Sheet.Shapes
        DoEvents
        Call UpdateStatusBar(Sheet.Parent.FullName)
        
        If HasTextFrameCharactersText(objShape) Then
            If objREG_.Test(objShape.TextFrame.Characters.text) Then
                '検索文字列を発見
                Call ProcessShape(objShape)
            End If
        End If
    Next
End Sub

'図形のテキストフレーム内に文字があるかを返す
Private Function HasTextFrameCharactersText(objShape As Shape) As Boolean
On Error Resume Next
    HasTextFrameCharactersText = (objShape.TextFrame.Characters.text <> "")
End Function

Private Sub ProcessCell(objCell As Range)
    Call CreateResultInfo(objSheet:=objCell.Parent, _
                          strAddress:=objCell.address, _
                          strValue:=objCell.value)
End Sub

Private Sub ProcessShape(objShape As Shape)
    Call CreateResultInfo(objSheet:=objShape.Parent, _
                          strAddress:=objShape.Name, _
                          strValue:=objShape.TextFrame.Characters.text)
End Sub

'検索結果の情報を作成する
Private Sub CreateResultInfo(objSheet As Worksheet, strAddress As String, strValue As String)
    Dim objRow As Range
    Dim objPathCell As Range
    Dim objBookCell As Range
    
    Set objRow = GetNewRow()
    Set objPathCell = objRow.Cells(ResultColumns.Path)
    Set objBookCell = objRow.Cells(ResultColumns.Book)
    
    'パス
    objPathCell.value = objSheet.Parent.FullName
    objPathCell.WrapText = False
    Call objMainSheet_.Hyperlinks.Add(Anchor:=objPathCell, address:=objPathCell.value)
    
    'ブック
    objBookCell.value = objSheet.Parent.Name
    Call objMainSheet_.Hyperlinks.Add(Anchor:=objBookCell, address:=objPathCell.value)
    
    'シート
    objRow.Cells(ResultColumns.Sheet).value = objSheet.Name
    
    '位置
    objRow.Cells(ResultColumns.Pos).value = strAddress
    
    'セルの値
    objRow.Cells(ResultColumns.Val).value = strValue
    objRow.Cells(ResultColumns.Val).WrapText = False
    
    '--- 罫線 ---
    objRow.Borders.LineStyle = xlContinuous
End Sub

Private Function GetNewRow() As Range
    Dim objListRange As Range
    Dim newRowIndex As Long
    
    Set objListRange = GetResultListRange
    newRowIndex = objListRange.Rows.Count + 1
    
    Set GetNewRow = objListRange.Rows(newRowIndex)
End Function

Private Function GetResultListRange() As Range
    Dim objListRowsCount As Long
    
    objListRowsCount = objResultListFirstPos_.CurrentRegion.Rows.Count - 1
    
    Set GetResultListRange = objResultListFirstPos_.CurrentRegion.Offset(1).Resize(objListRowsCount)
End Function

Private Sub UpdateStatusBar(msg As String)
    Const DOTSLEN As Integer = 5
    Dim dots As String
    Dim loading As String
    Dim display As String
    
    dots = String((Math.Rnd * DOTSLEN), ".")
    loading = Left(dots & String(DOTSLEN, " "), DOTSLEN)
    display = "検索中" & loading & " " & msg
    
    Application.StatusBar = IIf(msg = "", "", display)
    DoEvents
End Sub

Private Function IsIncludedInListRange(objTarget As Range) As Boolean
    Dim objResultRange As Range
    
    Set objResultRange = Application.Intersect(GetResultListRange, objTarget)
    
    IsIncludedInListRange = Not objResultRange Is Nothing
End Function

'検索結果のハイパーリンクをクリックしたときに発生するイベント
Public Sub FollowHyperlink(objTarget As Hyperlink)
    Dim objSourceCell As Range
    Dim objSourceSheet As Worksheet
    Dim objSourceRow As Range
    Dim sheetName As String
    Dim objectName As String
    Dim objBook As Workbook
    Dim isRangeObject As Boolean
    Dim objDistSheet As Worksheet
    
    If Not IsIncludedInListRange(objTarget.Range) Then
        Exit Sub
    End If
    
On Error Resume Next
    Set objSourceCell = objTarget.Range
    Set objSourceSheet = objSourceCell.Parent
    Set objSourceRow = objSourceSheet.Range(objSourceCell.End(xlToRight).End(xlToLeft), objSourceCell.End(xlToRight))
    sheetName = objSourceRow.Cells(ResultColumns.Sheet)
    objectName = objSourceRow.Cells(ResultColumns.Pos)
    Set objBook = ActiveSheet.Parent
    isRangeObject = (objectName Like "$*")
    Set objDistSheet = objBook.Sheets(sheetName)

    objDistSheet.Activate
    
    If isRangeObject Then
        objDistSheet.Range(objectName).Activate
    Else
        objDistSheet.Shapes(objectName).Select
    End If
End Sub
