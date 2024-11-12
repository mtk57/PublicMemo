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

    main_sheet.Range("D7").value = EncryptDecrypt(True, main_sheet.Range("D6").value)

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

Private Function EncryptDecrypt_OLD_VERSION(ByVal isEncrypt As Boolean, ByVal word As String) As String
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

Private Function EncryptDecrypt(ByVal isEncrypt As Boolean, ByVal word As String) As String
    Const NG_CHARS = " ""&*.;<=>|"  'パスワードに使用できない文字
    Dim i As Integer
    Dim currentChar As String
    Dim currentAsc As Integer
 
    EncryptDecrypt = ""
 
    'パスワードのバリデーション（暗号化時のみチェック）
    If isEncrypt Then
        ' 文字列が空の場合はエラー
        If Len(word) = 0 Then
            Err.Raise 53, , "String is empty."
        End If
        
        'ASCII範囲外文字とNG文字のチェック
        For i = 1 To Len(word)
            currentChar = Mid(word, i, 1)
            currentAsc = Asc(currentChar)
            
            'ASCII範囲外チェック（0-127以外はエラー）
            If currentAsc < 0 Or currentAsc > 127 Then
                Err.Raise 53, , "Out of ASCII code range."
            End If
            
            'NG文字チェック
            If InStr(NG_CHARS, currentChar) > 0 Then
                Err.Raise 53, , "The password contains characters that cannot be used..NG char=[" & currentChar & "]"
            End If
        Next i
    End If
 
    Dim ret As String
    ret = ""
    
    For i = 1 To Len(word)
        currentChar = Mid(word, i, 1)
        currentAsc = Asc(currentChar)
        
        If isEncrypt Then
            '暗号化処理
            If currentAsc >= &H20 And currentAsc <= &H2E Then
                '0x20-0x2Eの範囲の文字は0x70-0x7Eにマッピング
                ret = ret & Chr(currentAsc + &H50)  '0x50 = 80を加算
            Else
                '通常の暗号化（15を引く）
                ret = ret & Chr(currentAsc - 15)
            End If
        Else
            '復号化処理
            If currentAsc >= &H70 And currentAsc <= &H7E Then
                '0x70-0x7Eの範囲の文字は0x20-0x2Eに戻す
                ret = ret & Chr(currentAsc - &H50)  '0x50 = 80を減算
            Else
                '通常の復号化（15を足す）
                ret = ret & Chr(currentAsc + 15)
            End If
        End If
    Next i
    
    EncryptDecrypt = ret
End Function
