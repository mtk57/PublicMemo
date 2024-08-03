Attribute VB_Name = "Main"
Option Explicit

'<概要>
'指定された①配下のフォルダ(サブフォルダ含む)から拡張子「vbp」のファイルを探索し、発見した場合はvbpファイルのパスとvbpファイルの内容を新規シートに出力する。
'
'<詳細>
'・言語：Excel VBA
'・プログラムへの入力：
'　①フォルダパス(絶対パス)
'・出力シートの仕様：
'　(a) A2～Anセル：発見したvbpファイルの絶対パス
'　(b) B1～n1セル：vbpファイルの内容(キー名)
'　(c) B2～n2セル：vbpファイルの内容(キー名に対応する値)
'
'<動作例>
'・vbpファイルが以下のパスに存在するとする。
'　(#1)  C:\tmp\test1.vbp
'　(#2)  C:\tmp\sub\test2.vbp
'・それぞのvbpファイルの中身は以下とする。
'　(test1.vbp)
'　Type=Exe
'　Form=frmMain1.frm
'　Command32=""
'　Name="TestProject1"
'
'　(test2.vbp)
'　Type=Exe
'　Form=frmMain2.frm
'　ExeName32="TestProject2.exe"
'　Command32=""
'　Name="TestProject2"
'
'・この状態でExcel VBAマクロに「C:\tmp」を指定すると出力されるシートは以下となること。
'
'A2：C:\tmp\test1.vbp
'A3：C:\tmp\sub\test2.vbp
'B1：Type
'C1：Form
'D1：Command32
'E1：Name
'F1：ExeName32
'B2：Exe
'C2：frmMain1.frm
'D2：""
'E2："TestProject1"
'F2：""
'B3：Exe
'C3：frmMain2.frm
'D3：""
'E3："TestProject2"
'F3："TestProject2.exe"
'------

Public Sub Run_Click()
On Error GoTo ErrorHandler
    If Common.ShowYesNoMessageBox("VBPプロットを実行します") = False Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    main_sheet.Range("A3").value = "処理中..."

    Dim msg As String: msg = "正常に終了しました"

    If IsEnableDebugLog() = True Then
        Common.OpenLog ThisWorkbook.path + Application.PathSeparator + "VbpPlot.log"
    End If

    Common.WriteLog "------------------------------------"
    Common.WriteLog "★Start"

    Worksheets("main").Activate
    Process.Run

    Common.WriteLog "★End"
    GoTo FINISH
    
ErrorHandler:
    msg = "エラーが発生しました!" & vbCrLf & "Reason=" & Err.Description

FINISH:
    Common.WriteLog msg
    Common.CloseLog
    main_sheet.Range("A3").value = ""
    Application.DisplayAlerts = True
    MsgBox msg
End Sub

Private Function IsEnableDebugLog() As Boolean
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Const Clm = "O7"
    
    Dim is_debug_log_s As String: is_debug_log_s = main_sheet.Range(Clm).value
    
    If is_debug_log_s = "" Or _
       is_debug_log_s = "NO" Then
       IsEnableDebugLog = False
    Else
        IsEnableDebugLog = True
    End If
End Function

