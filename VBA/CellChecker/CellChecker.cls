VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objMainSheet_ As Worksheet

Private objTargetPathPos_ As Range
Private objTargetSheetPos_ As Range
Private objResultListFirstPos_ As Range

Private shouldSearch_ As Boolean

Private objFSO_ As New FileSystemObject
Private objExcelApp_ As Excel.Application

Private Enum ResultColumns
    Path = 1
    Book = 2
    Sheet = 3
    Cell1 = 4
    Cell2 = 5
    Cell3 = 6
    Cell4 = 7
    Cell5 = 8
    Cell6 = 9
End Enum

Private Sub Class_Initialize()
    Set objMainSheet_ = ThisWorkbook.Sheets("CellChecker")
    Set objTargetPathPos_ = objMainSheet_.Range("C3")
    Set objTargetSheetPos_ = objMainSheet_.Range("C4")
    Set objResultListFirstPos_ = objMainSheet_.Range("C17")
    shouldSearch_ = False
End Sub

Public Sub ExecSearch()
    Dim objFolder As Folder
    
    Set objFolder = GetTargetFolder()
    
    If objFolder Is Nothing Then
        ShowInfoMsgBox ("�����Ώۃt�H���_��������܂���ł����B")
        Exit Sub
    End If
    
    If Trim(objTargetSheetPos_.value) = "" Then
        ShowInfoMsgBox ("�V�[�g������͂��Ă��������B")
        Exit Sub
    End If
    
    If shouldSearch_ = True Then
        ShowInfoMsgBox ("�ʂ̌������������łɎ��s���ł��B")
        Exit Sub
    End If
    
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
    
    For Each objFile In objFolder.Files
        Select Case objFSO_.GetExtensionName(objFile.Path)
            Case "xlsx"
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
        If objSheet.Name = Trim(objTargetSheetPos_.value) Then
            Call SearchSheet(objSheet)
            GoTo BREAK_FOR
        End If
    Next
    
BREAK_FOR:
    Call objBook.Close(SaveChanges:=False)
End Sub

Private Sub SearchSheet(objSheet As Worksheet)
    Dim objCell As Range
    Dim objCell1 As Range
    Dim objCell2 As Range
    Dim objCell3 As Range
    Dim objCell4 As Range
    Dim objCell5 As Range
    Dim objCell6 As Range
    
    Dim strCell1 As String
    Dim strCell2 As String
    Dim strCell3 As String
    Dim strCell4 As String
    Dim strCell5 As String
    Dim strCell6 As String

    If shouldSearch_ = False Then
        Exit Sub
    End If
    
    Set objCell1 = objMainSheet_.Range("F17")
    Set objCell2 = objMainSheet_.Range("G17")
    Set objCell3 = objMainSheet_.Range("H17")
    Set objCell4 = objMainSheet_.Range("I17")
    Set objCell5 = objMainSheet_.Range("J17")
    Set objCell6 = objMainSheet_.Range("K17")
    
    strCell1 = objSheet.Range(objCell1.value).value
    strCell2 = objSheet.Range(objCell2.value).value
    strCell3 = objSheet.Range(objCell3.value).value
    strCell4 = objSheet.Range(objCell4.value).value
    strCell5 = objSheet.Range(objCell5.value).value
    strCell6 = objSheet.Range(objCell6.value).value

    Call CreateResultInfo(objSheet:=objSheet, _
                          strCell1:=strCell1, _
                          strCell2:=strCell2, _
                          strCell3:=strCell3, _
                          strCell4:=strCell4, _
                          strCell5:=strCell5, _
                          strCell6:=strCell6)

End Sub



'�������ʂ̏����쐬����
Private Sub CreateResultInfo(objSheet As Worksheet, strCell1 As String, strCell2 As String, strCell3 As String, strCell4 As String, strCell5 As String, strCell6 As String)
    Dim objRow As Range
    Dim objPathCell As Range
    Dim objBookCell As Range
    
    Set objRow = GetNewRow()
    Set objPathCell = objRow.Cells(ResultColumns.Path)
    Set objBookCell = objRow.Cells(ResultColumns.Book)
    
    '�p�X
    objPathCell.value = objSheet.Parent.FullName
    
    '�u�b�N
    objBookCell.value = objSheet.Parent.Name
    
    '�V�[�g
    objRow.Cells(ResultColumns.Sheet).value = objSheet.Name
    
    '�Z���̒l
    objRow.Cells(ResultColumns.Cell1).value = strCell1
    objRow.Cells(ResultColumns.Cell2).value = strCell2
    objRow.Cells(ResultColumns.Cell3).value = strCell3
    objRow.Cells(ResultColumns.Cell4).value = strCell4
    objRow.Cells(ResultColumns.Cell5).value = strCell5
    objRow.Cells(ResultColumns.Cell6).value = strCell6
    
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



