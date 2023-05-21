
hoge

fuga

'Sub Sub0()
'End Sub

Function Function0_1(ByVal arg As String) As String
WriteLogSimple "Function0_1 START"  'for DEBUG
WriteLogSimple "Function0_1 END 1"  'for DEBUG
	Exit Function

WriteLogSimple "Function0_1 END"  'for DEBUG
End Function
Sub Sub0_1(ByVal arg As String) As String
WriteLogSimple "Sub0_1 START"  'for DEBUG
WriteLogSimple "Sub0_1 END 1"  'for DEBUG
	Exit Sub

WriteLogSimple "Sub0_1 END"  'for DEBUG
End Function
Function Function0_2(ByVal arg As String) As String 'hoge
WriteLogSimple "Function0_2 START"  'for DEBUG
WriteLogSimple "Function0_2 END"  'for DEBUG
End Function
Function Function0_3(ByVal arg As String _
) As String 'hoge
WriteLogSimple "Function0_3 START"  'for DEBUG
WriteLogSimple "Function0_3 END"  'for DEBUG
End Function

Function Function0_4(ByVal arg As String _
) As String
WriteLogSimple "Function0_4 START"  'for DEBUG
WriteLogSimple "Function0_4 END"  'for DEBUG
End Function

Function Function0_5(ByVal arg As String _
) As String()
WriteLogSimple "Function0_5 START"  'for DEBUG
WriteLogSimple "Function0_5 END"  'for DEBUG
End Function

Function Function0_6(ByVal arg As String _
) As String()	'hoge
WriteLogSimple "Function0_6 START"  'for DEBUG
WriteLogSimple "Function0_6 END"  'for DEBUG
End Function

Function Function0_7(ByVal arg As String _
) As _
String()	'hoge
WriteLogSimple "Function0_7 START"  'for DEBUG
WriteLogSimple "Function0_7 END"  'for DEBUG
End Function

Function Function0_8(ByVal arg As String _
) As _
String()
WriteLogSimple "Function0_8 START"  'for DEBUG
WriteLogSimple "Function0_8 END"  'for DEBUG
End Function

Function Function0_9(ByVal arg As String _
)
WriteLogSimple "Function0_9 START"  'for DEBUG
WriteLogSimple "Function0_9 END"  'for DEBUG
End Function

Function Function0_10()
WriteLogSimple "Function0_10 START"  'for DEBUG
WriteLogSimple "Function0_10 END"  'for DEBUG
End Function
Function Function0_10( _
)
WriteLogSimple "Function0_10 START"  'for DEBUG
WriteLogSimple "Function0_10 END"  'for DEBUG
End Function

Function Function0_11(ByVal arg As String _
) _
As _
String()
WriteLogSimple "Function0_11 START"  'for DEBUG
WriteLogSimple "Function0_11 END"  'for DEBUG
End Function

Sub Sub0_1(ByVal arg As String)
WriteLogSimple "Sub0_1 START"  'for DEBUG
WriteLogSimple "Sub0_1 END"  'for DEBUG
End Sub
Sub Sub0_2(ByVal arg As String)	'hoge
WriteLogSimple "Sub0_2 START"  'for DEBUG
WriteLogSimple "Sub0_2 END"  'for DEBUG
End Sub
Sub Sub0_3(ByVal arg As String _
)
WriteLogSimple "Sub0_3 START"  'for DEBUG
WriteLogSimple "Sub0_3 END"  'for DEBUG
End Sub

Sub Sub0_4(ByVal arg As String _
)	'hoge
WriteLogSimple "Sub0_4 START"  'for DEBUG
WriteLogSimple "Sub0_4 END"  'for DEBUG
End Sub

Sub Sub0(ByVal arg As String. _
         ByRef arg2 As Object)  'hoge
WriteLogSimple "Sub0 START"  'for DEBUG

	Dim a As Long
	If True Then
		Exit Sub
	End If

WriteLogSimple "Sub0 END"  'for DEBUG
End Sub

Sub Sub1()
WriteLogSimple "Sub1 START"  'for DEBUG
WriteLogSimple "Sub1 END"  'for DEBUG
End Sub
Private Sub Sub2()
WriteLogSimple "Sub2 START"  'for DEBUG
WriteLogSimple "Sub2 END"  'for DEBUG
End Sub
Public Sub Sub3()
WriteLogSimple "Sub3 START"  'for DEBUG
WriteLogSimple "Sub3 END"  'for DEBUG
End Sub
Public Sub Sub4(ByVal arg As String)
WriteLogSimple "Sub4 START"  'for DEBUG
WriteLogSimple "Sub4 END"  'for DEBUG
End Sub
Public Sub Sub5(arg As String)
WriteLogSimple "Sub5 START"  'for DEBUG
WriteLogSimple "Sub5 END"  'for DEBUG
End Sub
Public Sub Sub6(arg)
WriteLogSimple "Sub6 START"  'for DEBUG
WriteLogSimple "Sub6 END"  'for DEBUG
End Sub


