VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContentsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ContentsManager

'メンバー変数
Private dict_infos_ As Object

'コンストラクタ
Private Sub Class_Initialize()
End Sub

'初期化
Public Sub Init(ByRef file_paths() As String)
    Set dict_infos_ = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim key As String
    Dim info As ContentsInfo
    
    For i = 0 To UBound(file_paths)
        key = file_paths(i)
    
        If Not dict_infos_.Exists(key) Then
            Set info = New ContentsInfo
            info.Init key
        
            dict_infos_.Add key, info
        End If
    Next i
End Sub

'=========================================
'Getter
'=========================================
Public Function GetContents(ByVal file_path As String) As ContentsInfo
    GetContents = Nothing
    
    If dict_infos_.Exists(file_path) Then
        GetContents = dict_infos_(file_path)
    End If
End Function


