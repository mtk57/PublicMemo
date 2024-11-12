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

    Dim msg As String: msg = "����ɏI�����܂���"


    Worksheets("main").Activate

    main_sheet.Range("D7").value = EncryptDecrypt(True, main_sheet.Range("D6").value)

    GoTo FINISH
    
ErrorHandler:
    msg = "�G���[���������܂���!" & vbCrLf & "Reason=" & Err.Description
    MsgBox msg

FINISH:
    Application.DisplayAlerts = True
End Sub

Public Sub Decrypt_Click()
On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    Dim msg As String: msg = "����ɏI�����܂���"


    Worksheets("main").Activate

    Dim ret As String
    ret = EncryptDecrypt(False, main_sheet.Range("D7").value)
    main_sheet.Range("D6").value = ret

    GoTo FINISH
    
ErrorHandler:
    msg = "�G���[���������܂���!" & vbCrLf & "Reason=" & Err.Description
    MsgBox msg

FINISH:
    Application.DisplayAlerts = True
End Sub

Private Function EncryptDecrypt_OLD_VERSION(ByVal isEncrypt As Boolean, ByVal word As String) As String
    Dim ret As String
    Dim i As Integer
    
    ret = ""
  
    If isEncrypt = True Then
        '�Í���
        For i = 1 To Len(word)
            ret = ret & Chr(Asc(Mid(word, i, 1)) - 15)
        Next i
    Else
        '������
        For i = 1 To Len(word)
            ret = ret & Chr(Asc(Mid(word, i, 1)) + 15)
        Next i
    End If

    EncryptDecrypt = ret

End Function

Private Function EncryptDecrypt(ByVal isEncrypt As Boolean, ByVal word As String) As String
    Const NG_CHARS = " ""&*.;<=>|"  '�p�X���[�h�Ɏg�p�ł��Ȃ�����
    Dim i As Integer
    Dim currentChar As String
    Dim currentAsc As Integer
 
    EncryptDecrypt = ""
 
    '�p�X���[�h�̃o���f�[�V�����i�Í������̂݃`�F�b�N�j
    If isEncrypt Then
        ' �����񂪋�̏ꍇ�̓G���[
        If Len(word) = 0 Then
            Err.Raise 53, , "String is empty."
        End If
        
        'ASCII�͈͊O������NG�����̃`�F�b�N
        For i = 1 To Len(word)
            currentChar = Mid(word, i, 1)
            currentAsc = Asc(currentChar)
            
            'ASCII�͈͊O�`�F�b�N�i0-127�ȊO�̓G���[�j
            If currentAsc < 0 Or currentAsc > 127 Then
                Err.Raise 53, , "Out of ASCII code range."
            End If
            
            'NG�����`�F�b�N
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
            '�Í�������
            If currentAsc >= &H20 And currentAsc <= &H2E Then
                '0x20-0x2E�͈̔͂̕�����0x70-0x7E�Ƀ}�b�s���O
                ret = ret & Chr(currentAsc + &H50)  '0x50 = 80�����Z
            Else
                '�ʏ�̈Í����i15�������j
                ret = ret & Chr(currentAsc - 15)
            End If
        Else
            '����������
            If currentAsc >= &H70 And currentAsc <= &H7E Then
                '0x70-0x7E�͈̔͂̕�����0x20-0x2E�ɖ߂�
                ret = ret & Chr(currentAsc - &H50)  '0x50 = 80�����Z
            Else
                '�ʏ�̕������i15�𑫂��j
                ret = ret & Chr(currentAsc + 15)
            End If
        End If
    Next i
    
    EncryptDecrypt = ret
End Function
