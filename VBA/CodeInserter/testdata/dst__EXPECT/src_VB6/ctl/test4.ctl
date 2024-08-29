hoge

fuga

Sub Sub1()
WriteLogSimple "test4.ctl" & vbTab & "Sub1" & vbTab & "START"  'for DEBUG
WriteLogSimple "test4.ctl" & vbTab & "Sub1" & vbTab & "END"  'for DEBUG
End Sub

Private Function Func2()
WriteLogSimple "test4.ctl" & vbTab & "Func2" & vbTab & "START"  'for DEBUG
WriteLogSimple "test4.ctl" & vbTab & "Func2" & vbTab & "END"  'for DEBUG
End Function

Public Sub Sub3()
WriteLogSimple "test4.ctl" & vbTab & "Sub3" & vbTab & "START"  'for DEBUG
WriteLogSimple "test4.ctl" & vbTab & "Sub3" & vbTab & "END"  'for DEBUG
End Sub

Public Function Func4(ByVal arg As String)
WriteLogSimple "test4.ctl" & vbTab & "Func4" & vbTab & "START"  'for DEBUG
WriteLogSimple "test4.ctl" & vbTab & "Func4" & vbTab & "END"  'for DEBUG
End Function

Public Sub Sub4(arg As String)
WriteLogSimple "test4.ctl" & vbTab & "Sub4" & vbTab & "START"  'for DEBUG
WriteLogSimple "test4.ctl" & vbTab & "Sub4" & vbTab & "END"  'for DEBUG
End Sub

Public Function Func5(arg)
WriteLogSimple "test4.ctl" & vbTab & "Func5" & vbTab & "START"  'for DEBUG
WriteLogSimple "test4.ctl" & vbTab & "Func5" & vbTab & "END"  'for DEBUG
End Function



