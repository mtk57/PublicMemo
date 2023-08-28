VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'メンバー変数
Private vbprj_file_paths() As String

Private Const CLM_ = "C"
Private Const START_ROW_ = 22

'コンストラクタ
Private Sub Class_Initialize()
    Erase vbprj_file_paths
End Sub

'初期化
Public Sub Init()
    CollectParam
End Sub

'初期化(外部実行用)
Public Sub InitForExternal(ByRef files() As String)
    vbprj_file_paths = files
End Sub

Private Sub CollectParam()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    Dim row As Integer: row = START_ROW_
    Dim i As Integer: i = 0
    Dim path As String
    
    Do
        path = main_sheet.Range(CLM_ & row + i).value
        
        If path = "" Then
            Exit Do
        End If
        
        ReDim Preserve vbprj_file_paths(i)
        vbprj_file_paths(i) = path
        
        i = i + 1
    Loop
End Sub

'検証
Public Sub Validate()
    Dim i As Integer
    Dim path As String
    Dim ext As String
    
    If GetVBProjFilePathListCount() <= 0 Then
        Exit Sub
    End If
    
    For i = LBound(vbprj_file_paths) To UBound(vbprj_file_paths)
        path = vbprj_file_paths(i)
        If Common.IsExistsFile(path) = False Then
            Err.Raise 53, , "[SubParam] VBプロジェクトファイルが見つかりません。" & vbCrLf & _
                            "path=" & path
        End If
        
        ext = Common.GetFileExtension(path)
        If ext <> "vbp" And ext <> "vbproj" Then
            Err.Raise 53, , "[SubParam] VBプロジェクトファイル名が未対応の拡張子です。" & vbCrLf & _
                            "path=" & path
        End If
    Next i

End Sub

'=========================================
'Getter
'=========================================

Public Function GetVBProjFilePathList() As String()
    GetVBProjFilePathList = vbprj_file_paths()
End Function

Public Function GetVBProjFilePathListCount() As Integer
    If Common.IsEmptyArray(vbprj_file_paths) = True Then
        GetVBProjFilePathListCount = 0
        Exit Function
    End If
    
    Dim cnt As Integer: cnt = UBound(vbprj_file_paths())
    cnt = cnt + 1
    GetVBProjFilePathListCount = cnt
End Function

