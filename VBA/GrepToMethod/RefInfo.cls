VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RefInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'RefInfo

'メンバー変数
Private proj_path_ As String
Private ref_paths_() As String

'コンストラクタ
Private Sub Class_Initialize()
    proj_path_ = ""
End Sub

'初期化
Public Sub Init()
End Sub

'=========================================
'Setter
'=========================================

Public Sub SetProjectPath(ByVal proj_path As String)
    proj_path_ = proj_path
End Sub

Public Sub AppendRefPath(ByVal ref_path As String)
    Common.AppendArray ref_paths_, ref_path
End Sub

'=========================================
'Getter
'=========================================

Public Function GetProjectPath() As String
    GetProjectPath = proj_path_
End Function

Public Function GetRefPaths() As String()
    GetRefPaths = ref_paths_
End Function

Public Function IsEmptyRefPaths() As Boolean
    IsEmptyRefPaths = Common.IsEmptyArray(ref_paths_)
End Function

Public Function GetRefPathsCount() As Long
    GetRefPathsCount = UBound(ref_paths_)
End Function

