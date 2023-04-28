Attribute VB_Name = "Process_002"
Option Explicit

Private prm As Param

Public Sub Run()
    Common.WriteLog "Run S"
    
    Dim msg As String: msg = ""

    Set prm = New Param
    
    prm.Init
    prm.Validate
    
    Common.WriteLog prm.GetAllValue()
    
    Clone
    
    Common.WriteLog "Run E"
End Sub

Private Sub Clone()
    Common.WriteLog "Clone S"

    
    Common.WriteLog "Clone E"
End Sub


