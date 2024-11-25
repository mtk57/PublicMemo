VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallMethodManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CallMethodManager

'メンバー変数

Private main_param_ As MainParam
Private infos_() As CallMethodInfo
Private refmgr_ As RefInfoManager

'コンストラクタ
Private Sub Class_Initialize()
End Sub

'初期化
Public Sub Init(ByRef main_param As MainParam, ByRef results() As GrepResultInfoStruct)
    Common.WriteLog "CallMethodManager.Init S"

    Set main_param_ = main_param

    If main_param.GetLang() = VB6 Then
        Call InitForVB6
    End If

    Common.WriteLog "CallMethodManager.Init E"
End Sub

Public Sub Parse()
    Common.WriteLog "CallMethodManager.Parse S"
    
    If main_param.GetLang() = VB6 Then
        Call ParseForVB6
    End If
    
    Common.WriteLog "CallMethodManager.Parse E"
End Sub

Public Sub OutputSheet(ByVal sheet_name As String)
    Common.WriteLog "CallMethodManager.OutputSheet S"
    
    If main_param.GetLang() = VB6 Then
        Call OutputSheetForVB6
    End If
    
    Common.WriteLog "CallMethodManager.OutputSheet E"
End Sub

Private Sub InitForVB6()
    Common.WriteLog "CallMethodManager.InitForVB6 S"

    Call CreateVbpRefList

    Dim i As Long
    Dim info_ As CallMethodInfo
    Dim result As GrepResultInfoStruct
    
    For i = 0 To UBound(results)
        result = results(i)
        
        ReDim Preserve infos_(i)
        
        Set info_ = New CallMethodInfo
        
        infos_(i) = info_
        
        info_.SetResultInfo result
        info_.SetLevel 0
        info_.SetParentPath result.FilePath
        
CONTINUE_I:
        
    Next i
    
    Common.WriteLog "CallMethodManager.InitForVB6 E"
End Sub

Private Sub CreateVbpRefList()
    Common.WriteLog "CallMethodManager.CreateVbpRefList S"
    
    Set refmgr_ = New RefInfoManager
    refmgr_.Init main_param_
    
    Common.WriteLog "CallMethodManager.CreateVbpRefList E"
End Sub

Private Sub ParseForVB6()
    Common.WriteLog "CallMethodManager.ParseForVB6 S"

    Common.WriteLog "CallMethodManager.ParseForVB6 E"
End Sub

Private Sub OutputSheetForVB6()
    Common.WriteLog "CallMethodManager.OutputSheetForVB6 S"
    
    Common.WriteLog "CallMethodManager.OutputSheetForVB6 E"
End Sub
