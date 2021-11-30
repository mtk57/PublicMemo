Attribute VB_Name = "MainModule"
Option Explicit

Public CC As New CellChecker

Public Sub 検索実行()
    Call CC.ExecSearch
End Sub

Public Sub 検索中止()
    Call CC.StopSearch
End Sub

Public Sub 結果リストをクリア()
    Call CC.ClearResultList
End Sub
