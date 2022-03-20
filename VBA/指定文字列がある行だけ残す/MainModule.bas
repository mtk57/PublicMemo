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

    Const FLUSH_BYTES = 1700
    Dim outputCount As Long
    Dim data As String
    
    outputCount = 0
    data = ""

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
            data = data & buffer & vbNewLine
        Else
            isFinding = False
        End If
        
        If FLUSH_BYTES <= LenB(data) Then
            If outputCount = 0 Then
                Call CreateOutputFile(outputFilePath, data)
            Else
                Call AppendOutputFile(outputFilePath, data)
            End If
            
            data = ""
            outputCount = outputCount + 1
        End If
        
    Wend

    If outputCount = 0 Then
        Call CreateOutputFile(outputFilePath, data)
    End If

FINISH:
    Close #1

    MsgBox msg
    
    Exit Sub

Exception:
    Close #1
    MsgBox Err.Number & vbCrLf & Err.Description

End Sub

'高速化
'参考：https://chuckischarles.hatenablog.com/entry/2018/03/24/211905

Sub CreateOutputFile(ByVal filePath As String, ByVal data As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    With fso.CreateTextFile(filePath)
        .WriteLine data
        .Close
    End With
    
    Set fso = Nothing
End Sub

Sub AppendOutputFile(ByVal filePath As String, ByVal data As String)
    Open filePath For Append As #2
    Print #2, data
    Close #2
End Sub
