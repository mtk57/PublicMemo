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
    
    If objIgnoreCasePos_ = "" Or objIgnoreCasePos_ = "��ʂ���i�f�t�H���g�j" Then
        ignoreCase = True
    End If
    
    If objFolder Is Nothing Then
        ShowInfoMsgBox ("�����Ώۃt�H���_��������܂���ł����B")
        Exit Sub
    End If
    
    If Trim(objSearchWordPos_.value) = "" Then
        ShowInfoMsgBox ("�������������͂��Ă��������B")
        Exit Sub
    End If
    
    If shouldSearch_ = True Then
        ShowInfoMsgBox ("�ʂ̌������������łɎ��s���ł��B")
        Exit Sub
    End If
    
    objREG_.Global = True
    objREG_.ignoreCase = ignoreCase
    objREG_.Pattern = objSearchWordPos_.value
    
    '�����J�n
    shouldSearch_ = True
    Call ClearResultList
    Call SearchFolder(objFolder)
    
    '�����I��
    If Not objExcelApp_ Is Nothing Then
        objExcelApp_.Quit
        Set objExcelApp_ = Nothing
    End If
    
    shouldSearch_ = False
    Call UpdateStatusBar("")
    ShowInfoMsgBox ("�������I�����܂����B")
End Sub

Public Sub StopSearch()
    If shouldSearch_ = False Then
        Exit Sub
    End If
    
    If ShowYNMsgBox("�����𒆎~���Ă���낵���ł���?") <> vbYes Then
        Exit Sub
    End If
    
    shouldSearch_ = False
End Sub

Public Sub ClearResultList()
    'CurrentRegion�̓A�N�e�B�u�Z���̈�i���j�̂��ƁB
    'Offset�̓A�N�e�B�u�Z���̈������2�s���炷�B
    '����Ō������ʂ̗̈悾�����ΏۂƂȂ�B�i�������ʂ�2�s��ɂ́u�\���v�Ɓu�\�̗�v��2�s����������Ă̓_���Ȃ̂Łj
    '
    '�� �A�N�e�B�u�Z���̈�Ƃ�
    '   �w�肵���Z�����܂ށA�󔒍s�Ƌ󔒗�ň͂܂ꂽ�A�f�[�^�̓��͔͈͂ł��B
    '   [CTRL] + [SHIFT] + [*] �ł��I���ł���
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
        ShowInfoMsgBox (objFile.Path & vbCrLf & " ���J���܂���ł����B")
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
                '����������𔭌�
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
                '����������𔭌�
                Call ProcessShape(objShape)
            End If
        End If
    Next
End Sub

'�}�`�̃e�L�X�g�t���[�����ɕ��������邩��Ԃ�
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

'�������ʂ̏����쐬����
Private Sub CreateResultInfo(objSheet As Worksheet, strAddress As String, strValue As String)
    Dim objRow As Range
    Dim objPathCell As Range
    Dim objBookCell As Range
    
    Set objRow = GetNewRow()
    Set objPathCell = objRow.Cells(ResultColumns.Path)
    Set objBookCell = objRow.Cells(ResultColumns.Book)
    
    '�p�X
    objPathCell.value = objSheet.Parent.FullName
    objPathCell.WrapText = False
    Call objMainSheet_.Hyperlinks.Add(Anchor:=objPathCell, address:=objPathCell.value)
    
    '�u�b�N
    objBookCell.value = objSheet.Parent.Name
    Call objMainSheet_.Hyperlinks.Add(Anchor:=objBookCell, address:=objPathCell.value)
    
    '�V�[�g
    objRow.Cells(ResultColumns.Sheet).value = objSheet.Name
    
    '�ʒu
    objRow.Cells(ResultColumns.Pos).value = strAddress
    
    '�Z���̒l
    objRow.Cells(ResultColumns.Val).value = strValue
    objRow.Cells(ResultColumns.Val).WrapText = False
    
    '--- �r�� ---
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
    display = "������" & loading & " " & msg
    
    Application.StatusBar = IIf(msg = "", "", display)
    DoEvents
End Sub

Private Function IsIncludedInListRange(objTarget As Range) As Boolean
    Dim objResultRange As Range
    
    Set objResultRange = Application.Intersect(GetResultListRange, objTarget)
    
    IsIncludedInListRange = Not objResultRange Is Nothing
End Function

'�������ʂ̃n�C�p�[�����N���N���b�N�����Ƃ��ɔ�������C�x���g
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