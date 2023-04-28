Attribute VB_Name = "Process_001"
Option Explicit

Private prms As ParamContainer

Public Sub Run()
    Common.WriteLog "Run S"
    
    Dim msg As String: msg = ""

    Set prms = New ParamContainer
    
    prms.Init
    prms.Validate
    
    Common.WriteLog prms.GetAllValue()
    
    '�O���c�[�����s
    Const MACRO_NAME As String = "Main.Run"
    Dim ret As Variant
    
    ret = Application.Run( _
          "'" & _
          prms.GetExternalPath() & _
          "'!" & _
          MACRO_NAME, _
          prms.GetVBProjFilePathList(), _
          prms.GetDestDirPath(), _
          prms.GetBaseFolder(), _
          prms.IsDebugLog() _
          )
    
    '�O���c�[�������
    Common.CloseBook (Common.GetFileName(prms.GetExternalPath()))
    
    If ret = False Then
        msg = "�O���c�[���̎��s�Ɏ��s���܂���!"
        Common.WriteLog "Run E2"
        Err.Raise 53, , msg
    End If
    
    Common.WriteLog "Run E"
End Sub


