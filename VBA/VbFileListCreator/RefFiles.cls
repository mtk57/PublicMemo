VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RefFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'メンバー変数

Private vbp_file_path As String
Private ref_file_path() As String


'コンストラクタ
Private Sub Class_Initialize()
    vbp_file_path = ""
    Erase ref_file_path
End Sub

'初期化
Public Sub Init()

End Sub

'=========================================
'Setter
'=========================================

'-----------------------------------------
'VBPファイルパス
'-----------------------------------------
Public Sub SetSrcDirPath(ByVal path As String)
    vbp_file_path = path
End Sub

'-----------------------------------------
'VBPが参照しているファイルパス
'-----------------------------------------
Public Sub AppendRefFilePath(ByVal path As String)
    Common.AppendArray ref_file_path, path
End Sub

'=========================================
'Getter
'=========================================

'-----------------------------------------
'VBPファイルパス
'-----------------------------------------
Public Function GetSrcDirPath() As String
    GetSrcDirPath = vbp_file_path
End Function

'-----------------------------------------
'VBPが参照しているファイルパスリストの要素数
'-----------------------------------------
Public Function GetAppendRefFileCount() As Long
    GetAppendRefFileCount = UBound(ref_file_path)
End Function

'-----------------------------------------
'VBPが参照しているファイルパスリスト
'-----------------------------------------
Public Function GetRefFile(ByVal index As Long) As String
    GetRefFile = ref_file_path(index)
End Function

'-----------------------------------------
'VBPが参照しているファイルパスリスト
'-----------------------------------------
Public Function GetRefFiles() As String()
    GetRefFiles = ref_file_path
End Function

