Attribute VB_Name = "CommonModule"
Option Explicit


Function ShowYNMsgBox(msg As String) As Integer
    ShowYNMsgBox = MsgBox(msg, vbYesNo Or vbQuestion, ThisWorkbook.Name)
End Function

Sub ShowInfoMsgBox(msg As String)
    MsgBox msg, vbInformation, ThisWorkbook.Name
End Sub

Function SelectFolderPath(msg As String) As String
    Dim objFileDialog As FileDialog
    Dim result As String
    
    Set objFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    objFileDialog.Title = msg
    
    If objFileDialog.Show Then
        result = objFileDialog.SelectedItems(1)
    End If
        
    SelectFolderPath = result
End Function

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
