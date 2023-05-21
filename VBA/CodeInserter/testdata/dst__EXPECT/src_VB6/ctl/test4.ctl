
hoge

fuga

Sub Sub1()
WriteLogSimple "Sub1 START"  'for DEBUG
WriteLogSimple "Sub1 END"  'for DEBUG
End Sub
Private Function Func2()
WriteLogSimple "Func2 START"  'for DEBUG
WriteLogSimple "Func2 END"  'for DEBUG
End Function
Public Sub Sub3()
WriteLogSimple "Sub3 START"  'for DEBUG
WriteLogSimple "Sub3 END"  'for DEBUG
End Sub
Public Function Func4(ByVal arg As String)
WriteLogSimple "Func4 START"  'for DEBUG
WriteLogSimple "Func4 END"  'for DEBUG
End Function
Public Sub Sub4(arg As String)
WriteLogSimple "Sub4 START"  'for DEBUG
WriteLogSimple "Sub4 END"  'for DEBUG
End Sub
Public Function Func5(arg)
WriteLogSimple "Func5 START"  'for DEBUG
WriteLogSimple "Func5 END"  'for DEBUG
End Function


