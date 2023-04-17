Attribute VB_Name = "Process_001"
Option Explicit

Private prm001 As Param_001

Public Sub Run()
    Common.WriteLog "Run S"
    
    Dim msg As String: msg = ""

    Set prm001 = New Param_001
    
    prm001.Init
    prm001.Validate
    
    Common.WriteLog prm001.GetAllValue()
    
    '�O���c�[�����s
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
    
    '�O���c�[�������
    Common.CloseBook (Common.GetFileName(prm001.GetExternalPath()))
    
    If ret = False Then
        msg = "�O���c�[���̎��s�Ɏ��s���܂���!"
        Common.WriteLog "Run E2"
        Err.Raise 53, , msg
    End If
    
    Common.WriteLog "Run E"
End Sub


