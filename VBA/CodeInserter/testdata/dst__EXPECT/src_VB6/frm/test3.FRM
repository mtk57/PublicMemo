hoge

fuga

	EditFunction = True

Sub Sub1()
WriteLogSimple "test3.FRM" & vbTab & "Sub1" & vbTab & "START"  'for DEBUG
WriteLogSimple "test3.FRM" & vbTab & "Sub1" & vbTab & "END"  'for DEBUG
End Sub

Private Sub Sub2()
WriteLogSimple "test3.FRM" & vbTab & "Sub2" & vbTab & "START"  'for DEBUG
WriteLogSimple "test3.FRM" & vbTab & "Sub2" & vbTab & "END"  'for DEBUG
End Sub

Public Sub Sub3()
WriteLogSimple "test3.FRM" & vbTab & "Sub3" & vbTab & "START"  'for DEBUG
WriteLogSimple "test3.FRM" & vbTab & "Sub3" & vbTab & "END"  'for DEBUG
End Sub

Public Sub Sub4(ByVal arg As String)
WriteLogSimple "test3.FRM" & vbTab & "Sub4" & vbTab & "START"  'for DEBUG
WriteLogSimple "test3.FRM" & vbTab & "Sub4" & vbTab & "END"  'for DEBUG
End Sub

Public Sub Sub5(arg As String)
WriteLogSimple "test3.FRM" & vbTab & "Sub5" & vbTab & "START"  'for DEBUG
WriteLogSimple "test3.FRM" & vbTab & "Sub5" & vbTab & "END"  'for DEBUG
End Sub

Public Sub Sub6(arg)
WriteLogSimple "test3.FRM" & vbTab & "Sub6" & vbTab & "START"  'for DEBUG
WriteLogSimple "test3.FRM" & vbTab & "Sub6" & vbTab & "END"  'for DEBUG
End Sub



