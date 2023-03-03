Attribute VB_Name = "CommonModule"
Option Explicit

'現在の日付と時間を文字列として取得する
Function GetTimeString() As String
    Dim strDate As String
    Dim strTime As String
    
    ' 現在の日付を取得し、文字列に変換する
    strDate = Format(Date, "yyyymmdd")
    
    ' 現在の時刻を取得し、文字列に変換する
    strTime = Format(Time, "hhmmss")
    
    ' 日付と時刻を結合して返す
    GetTimeString = strDate & strTime
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
