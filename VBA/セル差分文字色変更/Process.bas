Attribute VB_Name = "Process"
Option Explicit

'定数
Private SEP As String
Private DQ As String

Private Const SHEET_MAIN = "main"

Private Const SRC_CLM_ = "C"
Private Const DST_CLM_ = "D"
Private Const START_ROW_ = 10

'--------------------------------------------------------
'メイン処理
'--------------------------------------------------------
Public Sub Run()
    Common.WriteLog "Run S"
    
    Dim err_msg As String
    
    SEP = Application.PathSeparator
    DQ = Chr(34)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MAIN)

    Dim row As Integer: row = START_ROW_
    Dim i As Integer: i = 0
    Dim srcAdr As String
    Dim dstAdr As String
    Dim srcVal As String
    Dim dstVal As String
    
    Do
        srcAdr = SRC_CLM_ & row + i
        dstAdr = DST_CLM_ & row + i
    
        srcVal = ws.Range(srcAdr).value
        dstVal = ws.Range(dstAdr).value
        
        If srcVal = "" And dstVal = "" Then
            '空セル検知なので終了
            Exit Do
        End If
        
        If srcVal = dstVal Then
            Common.WriteLog "[" & i & "] SKIP  srcAdr=[" & srcAdr & "], srcVal=[" & srcVal & "], dstAdr=[" & dstAdr & "]. dstVal=[" & dstVal & "]"
        
            '差異が無いので無視
            GoTo CONTINUE_I
        End If
        
        Common.WriteLog "[" & i & "] srcAdr=[" & srcAdr & "], srcVal=[" & srcVal & "], dstAdr=[" & dstAdr & "]. dstVal=[" & dstVal & "]"
        
        '差異がある
        Call Common.CompareCellsAndHighlight(SHEET_MAIN, srcAdr, dstAdr)

        
CONTINUE_I:
        i = i + 1
    Loop
    
    Common.WriteLog "Run E"
End Sub
