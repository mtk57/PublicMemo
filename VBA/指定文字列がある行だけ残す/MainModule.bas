Attribute VB_Name = "MainModule"
Option Explicit

Sub ボタン1_Click()

On Error GoTo Exception
    Dim msg As String
    Dim inputFilePath As String
    Dim outputFilePath As String
    Dim keyword As String
    Dim first_keyword As String
    Dim buffer As String
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    Dim isFindFirstKeyword As Boolean
    Dim isFindKeyword As Boolean
    Dim isFinding As Boolean

    msg = "success"

    Worksheets("main").Select
                
    'パラメータ取得
    inputFilePath = Range("B5").Value
    outputFilePath = Range("B8").Value
    keyword = Range("B11").Value
    first_keyword = Range("B14").Value

    'パラメータチェック
    If dir(inputFilePath) = "" Then
        msg = "input file is not exist"
        GoTo FINISH
    End If

    If outputFilePath = "" Then
        msg = "output file is nothing"
        GoTo FINISH
    End If

    If keyword = "" Then
        msg = "keyword is nothing"
        GoTo FINISH
    End If

    If first_keyword = "" Then
        msg = "first keyword is nothing"
        GoTo FINISH
    End If
    
    
    '---------------------------------------
    'メイン

    Open inputFilePath For Input As #1
    Open outputFilePath For Output As #2
    
    reg.Pattern = first_keyword
    reg.Global = True
    
    isFinding = False
    
    While Not EOF(1)
        Line Input #1, buffer

        isFindKeyword = False
        isFindFirstKeyword = False

        If InStr(buffer, keyword) > 0 Then
            isFindKeyword = True
        End If
        
        isFindFirstKeyword = reg.Test(buffer)

        If (isFinding = False And isFindKeyword = True And isFindFirstKeyword = True) Or _
           (isFinding = True And (isFindKeyword = True And isFindFirstKeyword = True)) Or _
           (isFinding = True And (isFindKeyword = True And isFindFirstKeyword = False)) Or _
           (isFinding = True And (isFindKeyword = False And isFindFirstKeyword = False)) Then
            isFinding = True
            Print #2, buffer
        Else
            isFinding = False
        End If
        
    Wend


FINISH:
    Close #1
    Close #2

    MsgBox msg
    
    Exit Sub

Exception:
    Close #1
    Close #2
    MsgBox Err.Number & vbCrLf & Err.Description

End Sub



