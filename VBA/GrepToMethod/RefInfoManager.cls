VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RefInfoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'RefInfoManager

'メンバー変数

Private main_param_ As MainParam
Private infos_() As RefInfo

Private Const CLM_KEY = "A"
Private Const CLM_VALUE = "B"
Private Const ROW_START = 1

'コンストラクタ
Private Sub Class_Initialize()
End Sub

'初期化
Public Sub Init(ByRef main_param As MainParam)
    Common.WriteLog "RefInfoManager.Init S"

    Set main_param_ = main_param

    Dim target_sheet As Worksheet
    Set target_sheet = ThisWorkbook.Sheets(main_param_.GetVbpRefListSheetName())

    Dim row As Long: row = ROW_START
    Dim i As Long: i = 0
    
    Dim key_data As String
    Dim value_data As String
    Dim before_key As String
    
    Dim info As RefInfo
    Dim prj_cnt As Long: prj_cnt = 0
    
    Do
        key_data = target_sheet.Range(CLM_KEY & row + i).value
        value_data = target_sheet.Range(CLM_VALUE & row + i).value
        
        If key_data = "" Then
            'リストの終端を検知
            Exit Do
        End If
        
        If before_key <> key_data Then
            If Common.IsExistsFile(key_data) = False Then
                Set target_sheet = Nothing
                Err.Raise 53, , "VBP参照ファイル一覧のシートのA列のファイルが存在しません。(" & key_data & ")"
            End If
        End If
        
        If value_data = "" Then
            Set target_sheet = Nothing
            Err.Raise 53, , "VBP参照ファイル一覧のシートのB列にデータが存在しません。(i=" & i & ")"
        ElseIf Common.IsExistsFile(value_data) = False Then
            Set target_sheet = Nothing
            Err.Raise 53, , "VBP参照ファイル一覧のシートのB列のファイルが存在しません。(" & value_data & ")"
        End If
                
        If before_key <> key_data Then
            before_key = key_data
            
            Set info = New RefInfo
            
            ReDim Preserve infos_(prj_cnt)
            Set infos_(prj_cnt) = info
            
            info.SetProjectPath key_data
            info.AppendRefPath value_data
            
            prj_cnt = prj_cnt + 1
            
        ElseIf before_key = key_data Then
            info.AppendRefPath value_data
        End If
        
CONTINUE:
        i = i + 1
    Loop
    
    Set target_sheet = Nothing

    Common.WriteLog "RefInfoManager.Init E"
End Sub

'引数の紐づくプロジェクトファイルのパスを取得する
Public Function GetProjPaths(ByVal ref_path As String) As String()
    Dim i As Long
    Dim j As Long
    
    'TODO


End Function
