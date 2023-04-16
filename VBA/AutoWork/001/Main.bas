Attribute VB_Name = "Main"
Option Explicit

Private prm001 As Param001

Public Sub Run001_Click()
On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    
    Dim msg As String: msg = ""
    
    Set prm001 = New Param001
    prm001.Init
    msg = prm001.Validate()
    If msg <> "" Then
        GoTo FINISH
    End If
    
    '外部ツール実行
    Const MACRO_NAME As String = "Main.Run"
    Dim ret As Variant
    
    ret = Application.Run( _
          "'" & _
          prm001.GetExternalPath() & _
          "'!" & _
          MACRO_NAME, _
          prm001.GetVBProjFilePathList(), _
          prm001.GetDestDirPath(), _
          prm001.IsDebugLog() _
          )
    
    If ret = False Then
        msg = "外部ツールの実行に失敗しました!"
        GoTo FINISH
    End If

    msg = "正常に終了しました"
    GoTo FINISH

ErrorHandler:
    msg = "エラーが発生しました(" & Err.Description & ")"

FINISH:
    MsgBox msg
    Application.DisplayAlerts = True
End Sub
