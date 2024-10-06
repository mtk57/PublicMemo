Attribute VB_Name = "Process"
Option Explicit

'�萔
Private SEP As String
Private DQ As String

Private Const SHEET_MAIN = "main"

Private Const SRC_CLM_ = "C"
Private Const DST_CLM_ = "D"
Private Const START_ROW_ = 10

'--------------------------------------------------------
'���C������
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
            '��Z�����m�Ȃ̂ŏI��
            Exit Do
        End If
        
        If srcVal = dstVal Then
            Common.WriteLog "[" & i & "] SKIP  srcAdr=[" & srcAdr & "], srcVal=[" & srcVal & "], dstAdr=[" & dstAdr & "]. dstVal=[" & dstVal & "]"
        
            '���ق������̂Ŗ���
            GoTo CONTINUE_I
        End If
        
        Common.WriteLog "[" & i & "] srcAdr=[" & srcAdr & "], srcVal=[" & srcVal & "], dstAdr=[" & dstAdr & "]. dstVal=[" & dstVal & "]"
        
        '���ق�����
        Call Common.CompareCellsAndHighlight(SHEET_MAIN, srcAdr, dstAdr)

        
CONTINUE_I:
        i = i + 1
    Loop
    
    Common.WriteLog "Run E"
End Sub
