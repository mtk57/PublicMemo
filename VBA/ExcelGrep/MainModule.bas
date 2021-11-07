Attribute VB_Name = "MainModule"
Option Explicit

Public EG As New ExcelGrep

Public Sub 検索実行()
    Call EG.ExecSearch
End Sub

Public Sub 検索中止()
    Call EG.StopSearch
End Sub

Public Sub 結果リストをクリア()
    Call EG.ClearResultList
End Sub
