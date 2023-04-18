Attribute VB_Name = "Process_001"
Option Explicit

Private prm As Param

Public Sub Run()
    Common.WriteLog "Run S"
    
    Dim msg As String: msg = ""

    Set prm = New Param
    
    prm.Init
    prm.Validate
    
    Common.WriteLog prm.GetAllValue()
    
    '�O���c�[�����s
    Const MACRO_NAME As String = "Main.Run"
    Dim ret As Variant
    
    ret = Application.Run( _
          "'" & _
          prm.GetExternalPath() & _
          "'!" & _
          MACRO_NAME, _
          prm.GetVBProjFilePathList(), _
          prm.GetDestDirPath(), _
          prm.IsDebugLog() _
          )
    
    '�O���c�[�������
    Common.CloseBook (Common.GetFileName(prm.GetExternalPath()))
    
    If ret = False Then
        msg = "�O���c�[���̎��s�Ɏ��s���܂���!"
        Common.WriteLog "Run E2"
        Err.Raise 53, , msg
    End If
    
    Common.WriteLog "Run E"
End Sub


