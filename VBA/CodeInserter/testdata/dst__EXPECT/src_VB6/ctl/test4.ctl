hoge

fuga

Sub Sub1()
WriteLogSimple "test4.ctl:Sub1 START"  'for DEBUG
WriteLogSimple "test4.ctl:Sub1 END"  'for DEBUG
End Sub

Private Function Func2()
WriteLogSimple "test4.ctl:Func2 START"  'for DEBUG
WriteLogSimple "test4.ctl:Func2 END"  'for DEBUG
End Function

Public Sub Sub3()
WriteLogSimple "test4.ctl:Sub3 START"  'for DEBUG
WriteLogSimple "test4.ctl:Sub3 END"  'for DEBUG
End Sub

Public Function Func4(ByVal arg As String)
WriteLogSimple "test4.ctl:Func4 START"  'for DEBUG
WriteLogSimple "test4.ctl:Func4 END"  'for DEBUG
End Function

Public Sub Sub4(arg As String)
WriteLogSimple "test4.ctl:Sub4 START"  'for DEBUG
WriteLogSimple "test4.ctl:Sub4 END"  'for DEBUG
End Sub

Public Function Func5(arg)
WriteLogSimple "test4.ctl:Func5 START"  'for DEBUG
WriteLogSimple "test4.ctl:Func5 END"  'for DEBUG
End Function



