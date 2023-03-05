Attribute VB_Name = "MainModule"
Option Explicit

Const MAIN_SHEET = "MAIN"
Const FOLDER_PATH = "C4"
Const OUT_SHEET_NAME = "C6"
Const START_CELL_ROW = 1
Const START_CELL_CLM = 1
Const TOP_ROW = 1

Dim now_row As Integer
Dim replace_path As String
Dim last_clm As Integer

Sub CreateFileList()
    On Error GoTo ErrorHandler

    Dim inputFolderPath As String
    Dim outputSheetName As String
    Dim fso As Object
    
    now_row = 0
    last_clm = 0
    
    Application.DisplayAlerts = False
    
    Worksheets(MAIN_SHEET).Activate
    inputFolderPath = Range(FOLDER_PATH).Value
    outputSheetName = Range(OUT_SHEET_NAME).Value
    
    If outputSheetName = "" Then
        outputSheetName = getTimeString()
    End If
    AddSheet outputSheetName

    replace_path = removeLastFolder(inputFolderPath)
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    '�S�t�@�C�����������\�ɂ���
    Call allFiles(inputFolderPath, fso)
    
    '�\�����₷������
    Call formatTable
    
    MsgBox "�I��!"
    
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "�G���[���������܂����F" & Err.Description, vbCritical, "�G���["
    Application.DisplayAlerts = True
End Sub

'�S�t�@�C�����������\�ɂ���
Private Function allFiles(ByVal inputFolderPath As String, ByVal fso As Object)
    Dim folder As Object
    Dim file As Object
    Dim newPath As String
    
    For Each folder In fso.getFolder(inputFolderPath).SubFolders
        Call allFiles(folder.path, fso)
    Next

    For Each file In fso.getFolder(inputFolderPath).Files
        newPath = getNewFolderPath(file.path)
        splitPathToCells newPath
        now_row = now_row + 1
    Next
End Function

'�\�����₷������
Private Sub formatTable()
    Dim cell As Object

    last_clm = getLastColumnNumber()
    
    For Each cell In Range("A1").CurrentRegion
        'cell.Activate   'for DEBUG
        
        If cell.Column >= last_clm Then
            'To Next Row
        Else
            setGrayColorIfSameAsAbove cell
            moveCellValueToColumn cell, last_clm
        End If

    Next
    
End Sub

'�Z���̒l���w�肵���J�����ʒu�̃Z���Ɉړ�����
Private Sub moveCellValueToColumn(cell As Range, clmNum As Integer)
    If isCellEmpty(cell) = False And _
       isRightCellEmpty(cell) = True And _
       isLastColumn(cell) = False Then
        cell.Copy Destination:=cell.Offset(0, clmNum - cell.Column)
        cell.ClearContents
    End If
End Sub

'�ŏI�񂩂ǂ�����Ԃ�
Private Function isLastColumn(cell As Range) As Boolean
    If cell.Column >= last_clm Then
        isLastColumn = True
    Else
        isLastColumn = False
    End If
End Function

'�E�ɋ�Z�������邩��Ԃ�
Private Function isRightCellEmpty(cell As Range) As Boolean
    If IsEmpty(cell.Offset(0, 1)) Then
        isRightCellEmpty = True
    Else
        isRightCellEmpty = False
    End If
End Function

'��Z�����ǂ�����Ԃ�
Private Function isCellEmpty(cell As Range) As Boolean
    If cell.Value = "" Then
        isCellEmpty = True
    Else
        isCellEmpty = False
    End If
End Function

'�ŏI��ԍ���Ԃ�
Private Function getLastColumnNumber() As Integer
    Dim lastClm As Integer
    lastClm = Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    getLastColumnNumber = lastClm
End Function

'��΃p�X����A�w��t�H���_�ȍ~�̃p�X������Ԃ�
Private Function getNewFolderPath(ByVal path As String) As String
    Dim newPath As String
    newPath = Replace(path, replace_path, "")
    getNewFolderPath = newPath
End Function

'�p�X��\�ŕ������ă��[�N�V�[�g�̃Z���ɏo�͂���
Private Sub splitPathToCells(ByVal path As String)
    Dim pathParts As Variant
    pathParts = Split(path, "\")
    Dim i As Integer
    
    For i = 0 To UBound(pathParts)
        Cells(START_CELL_ROW + now_row, i + START_CELL_CLM) = pathParts(i)
    Next i
End Sub

'
'��F"C:\abc\def\xyz"�̏ꍇ�A"C:\abc\def\"���Ԃ�B
Private Function removeLastFolder(ByVal path As String) As String
    Dim lastIndex As Long
    Dim newPath As String
    
    lastIndex = InStrRev(path, "\")
    
    If lastIndex > 0 Then
        newPath = Left(path, lastIndex)
    Else
        newPath = path
    End If
    
    removeLastFolder = newPath
End Function

'1��̃Z���Ɠ����l�ł���΃t�H���g�F���O���[�ɂ���
Private Sub setGrayColorIfSameAsAbove(cell As Range)
    If cell.row = TOP_ROW Then
        Exit Sub
    End If

    If cell.Value = cell.Offset(-1, 0).Value Then
        cell.Font.Color = RGB(192, 192, 192)
    End If
End Sub

'���ݓ����𕶎���ŕԂ�
Private Function getTimeString() As String
    Dim strDate As String
    Dim strTime As String
    
    strDate = Format(Date, "yyyymmdd")
    strTime = Format(Time, "hhmmss")
    
    getTimeString = strDate & strTime
End Function

Private Function IsExistSheet(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.Name = sheetName Then
            IsExistSheet = True
            Exit Function
        End If
    Next ws
    
    IsExistSheet = False
End Function

Private Function AddSheet(ByVal sheetName As String)
    If IsExistSheet(sheetName) = True Then
        Application.DisplayAlerts = False
        Sheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If

    Worksheets.Add.Name = sheetName
End Function

