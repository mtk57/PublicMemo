Attribute VB_Name = "Process_Work01_001"
Option Explicit

Private prm001 As Param_Work01_001

Public Sub Run()
    Common.WriteLog "Run S"
    
    Dim msg As String: msg = ""

    Set prm001 = New Param_Work01_001
    
    prm001.Init
    
    msg = prm001.Validate()
    If msg <> "" Then
        Common.WriteLog "Run E1"
        Err.Raise 53, , msg
    End If
    
    Common.WriteLog prm001.GetAllValue()
    
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
        Common.WriteLog "Run E2"
        Err.Raise 53, , msg
    End If
    
    Common.WriteLog "Run E"
End Sub


