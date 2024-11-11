Attribute VB_Name = "Main"
Option Explicit

Public Sub Clear_Click()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    main_sheet.Range("D6").value = ""
    main_sheet.Range("D7").value = ""
End Sub


Public Sub Encrypt_Click()
On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    Dim msg As String: msg = "正常に終了しました"


    Worksheets("main").Activate

    Dim ret As String
    ret = EncryptDecrypt(True, main_sheet.Range("D6").value)
    main_sheet.Range("D7").value = ret

    GoTo FINISH
    
ErrorHandler:
    msg = "エラーが発生しました!" & vbCrLf & "Reason=" & Err.Description
    MsgBox msg

FINISH:
    Application.DisplayAlerts = True
End Sub

Public Sub Decrypt_Click()
On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    Dim msg As String: msg = "正常に終了しました"


    Worksheets("main").Activate

    Dim ret As String
    ret = EncryptDecrypt(False, main_sheet.Range("D7").value)
    main_sheet.Range("D6").value = ret

    GoTo FINISH
    
ErrorHandler:
    msg = "エラーが発生しました!" & vbCrLf & "Reason=" & Err.Description
    MsgBox msg

FINISH:
    Application.DisplayAlerts = True
End Sub

Private Function EncryptDecrypt(ByVal isEncrypt As Boolean, ByVal word As String) As String
    Dim ret As String
    Dim i As Integer
    
    ret = ""
  
    If isEncrypt = True Then
        '暗号化
        For i = 1 To Len(word)
            ret = ret & Chr(Asc(Mid(word, i, 1)) - 15)
        Next i
    Else
        '復号化
        For i = 1 To Len(word)
            ret = ret & Chr(Asc(Mid(word, i, 1)) + 15)
        Next i
    End If

    EncryptDecrypt = ret

End Function
