VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallMethodInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CallMethodInfo

'メンバー変数

Private resultInfo_ As GrepResultInfoStruct
Private level_ As Long
Private parent_fullpath_ As String
Private fullpath_ As String
Private signature_ As String
Private code_ As String
Private proj_paths_() As String

'コンストラクタ
Private Sub Class_Initialize()
    level_ = -1
    parent_fullpath_ = ""
    fullpath_ = ""
    signature_ = ""
    code_ = ""
End Sub

'初期化
Public Sub Init()

End Sub

'=========================================
'Setter
'=========================================

Public Sub SetResultInfo(ByRef result As GrepResultInfoStruct)
    resultInfo_ = result
End Sub

Public Sub SetLevel(ByVal level As Long)
    level_ = level
End Sub

Public Sub SetParentPath(ByVal parent_fullpath As String)
    parent_fullpath_ = parent_fullpath
End Sub

Public Sub SetPath(ByVal fullpath As String)
    fullpath_ = fullpath
End Sub

Public Sub SetSignature(ByVal signature As String)
    signature_ = signature
End Sub

Public Sub SetCode(ByVal code As String)
    code_ = code
End Sub

Public Sub AppendProjPath(ByVal proj_path As String)
    Common.AppendArray proj_paths_, proj_path
End Sub


'=========================================
'Getter
'=========================================

Public Function GetResultInfo() As GrepResultInfoStruct
    GetResultInfo = resultInfo_
End Function

Public Function GetLevel() As Long
    GetLevel = level_
End Function

Public Function GetParentPath() As String
    GetParentPath = parent_fullpath_
End Function

Public Function GetPath() As String
    GetPath = fullpath_
End Function

Public Function GetSignature() As String
    GetSignature = signature_
End Function

Public Function GetCode() As String
    GetCode = code_
End Function

Public Function GetProjPaths() As String()
    GetProjPaths = proj_paths_
End Function

Public Function IsEmptyProjPaths() As Boolean
    IsEmptyProjPaths = Common.IsEmptyArray(proj_paths_)
End Function

Public Function GetProjPathsCount() As Long
    GetProjPathsCount = UBound(proj_paths_)
End Function
