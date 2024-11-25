VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContentsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ContentsInfo

'メンバー変数
Private file_path_ As String
Private contents_() As String

'コンストラクタ
Private Sub Class_Initialize()
    file_path_ = ""
End Sub

'初期化
Public Sub Init(ByVal file_path As String)
    file_path_ = file_path
    
    Call ReadContents
End Sub

Private Sub ReadContents()
    contents_ = Common.GetContents(file_path_)
End Sub

'=========================================
'Getter
'=========================================

Public Function GetFilePath() As String
    GetFilePath = file_path_
End Function

Public Function GetContents() As String()
    GetContents = contents_
End Function

Public Function IsEmptyContents() As Boolean
    IsEmptyContents = Common.IsEmptyArray(contents_)
End Function

Public Function GetContentsLength() As Long
    GetContentsLength = UBound(contents_)
End Function

