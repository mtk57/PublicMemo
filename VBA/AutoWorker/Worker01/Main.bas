Attribute VB_Name = "Main"
Option Explicit

Private Const RUN_001 = "RUN_001"
Private Const RUN_002 = "RUN_002"

Sub Sample2()
    Dim objShell As Object
    Dim objExec As Object
    Dim strCommand As String
    Dim strOutput As String
    Dim objFSO As Object
    Dim objTextFile As Object
    Dim strFilePath As String

    ' Set the command you want to run
    'strCommand = Chr(34) & "C:\Program Files\Git\usr\bin\bash.exe" & Chr(34) & " --login -i -c & cd C:\_tmp\GIT_0020 & git log --oneline > C:\_tmp\test.txt"
    strCommand = Chr(34) & "C:\Program Files\Git\usr\bin\bash.exe" & Chr(34) & " --login -i -c & cd C:\_tmp\GIT_0020 & git log --oneline > C:\Users\motok\AppData\Local\Temp\20230428160434.txt"
    
    'Dim utf8_ As String: utf8_ = Common.ReadTextFileByUTF8("C:\_tmp\test.txt")

    ' Create a Shell object
    Set objShell = CreateObject("WScript.Shell")

    ' Execute the command and capture the output
    Set objExec = objShell.Exec("cmd.exe /c " & strCommand)

    ' Read the output
    'strOutput = objExec.stdout.ReadAll

    ' Set the path for the output text file
    'strFilePath = "C:\_tmp\output.txt"

    ' Create a FileSystemObject
    'Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Create a new text file and write the output to it
    'Set objTextFile = objFSO.CreateTextFile(strFilePath, True)
    'objTextFile.Write strOutput
    'objTextFile.Close

    ' Clean up
    'Set objTextFile = Nothing
    'Set objFSO = Nothing
    Set objExec = Nothing
    Set objShell = Nothing
End Sub



Public Sub Run001_Click()
On Error GoTo ErrorHandler
    'Sample2
    
    Dim result() As String
    result = Common.RunGit("C:\_tmp\GIT_0020", "git log --oneline")
    
    Exit Sub
    
    If Common.ShowYesNoMessageBox("[001]を実行します") = False Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    
    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "AutoRun_" & RUN_001 & ".log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets(Define.SHEET_01).Activate
    Process_001.Run

    Common.WriteLog "★End"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    MsgBox msg
End Sub

Public Sub Run002_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("[002]を実行します") = False Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    
    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "AutoRun_" & RUN_002 & ".log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets(Define.SHEET_01).Activate
    Process_002.Run

    Common.WriteLog "★End"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    Application.DisplayAlerts = True
    MsgBox msg
End Sub

Private Function IsEnableDebugLog() As Boolean
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(Define.SHEET_01)
    
    Dim is_debug_log_s As String: is_debug_log_s = main_sheet.Range(Define.DEBUG_LOG_CELL).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function

