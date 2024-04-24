Attribute VB_Name = "CommonModule"
Option Explicit

Function IsExistSheet(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.Name = sheetName Then
            IsExistSheet = True
            Exit Function
        End If
    Next ws
    
    IsExistSheet = False
End Function

Function AddSheet(ByVal sheetName As String)
    If IsExistSheet(sheetName) = True Then
        Application.DisplayAlerts = False
        Sheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If

    Worksheets.Add.Name = sheetName
End Function

Public Function ShowYesNoMessageBox(ByVal msg As String) As Boolean
    Dim result As Integer: result = MsgBox(msg, vbYesNo, "Confirm")
    
    If result = vbYes Then
        ShowYesNoMessageBox = True
    Else
        ShowYesNoMessageBox = False
    End If
End Function

'指定されたセルから列内の最後の使用済みセルまでをクリアする
Public Sub ClearRange(ByVal cell_address As String)
    On Error Resume Next
    Dim ws As Worksheet
    Dim last_row As Long
    Dim range_to_clear As Range
    
    Set ws = ActiveSheet
    
    ' 指定されたセルから列内の最後の使用済みセルまでの範囲を取得
    last_row = ws.Cells(ws.Rows.Count, ws.Range(cell_address).Column).End(xlUp).row
    Set range_to_clear = ws.Range(cell_address, ws.Cells(last_row, ws.Range(cell_address).Column))
    
    ' 範囲をクリア
    'range_to_clear.Clear
    range_to_clear.Value = ""
End Sub
