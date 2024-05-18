hoge

fuga

	EditFunction = True

Sub Sub1()
WriteLogSimple "test3.FRM:Sub1 START"  'for DEBUG
WriteLogSimple "test3.FRM:Sub1 END"  'for DEBUG
End Sub

Private Sub Sub2()
WriteLogSimple "test3.FRM:Sub2 START"  'for DEBUG
WriteLogSimple "test3.FRM:Sub2 END"  'for DEBUG
End Sub

Public Sub Sub3()
WriteLogSimple "test3.FRM:Sub3 START"  'for DEBUG
WriteLogSimple "test3.FRM:Sub3 END"  'for DEBUG
End Sub

Public Sub Sub4(ByVal arg As String)
WriteLogSimple "test3.FRM:Sub4 START"  'for DEBUG
WriteLogSimple "test3.FRM:Sub4 END"  'for DEBUG
End Sub

Public Sub Sub5(arg As String)
WriteLogSimple "test3.FRM:Sub5 START"  'for DEBUG
WriteLogSimple "test3.FRM:Sub5 END"  'for DEBUG
End Sub

Public Sub Sub6(arg)
WriteLogSimple "test3.FRM:Sub6 START"  'for DEBUG
WriteLogSimple "test3.FRM:Sub6 END"  'for DEBUG
End Sub



