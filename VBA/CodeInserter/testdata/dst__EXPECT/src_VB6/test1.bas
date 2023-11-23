Private Declare PtrSafe Function GetPrivateProfileString Lib _
    "kernel32" Alias "GetPrivateProfileStringA" ( _
) As Long

Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
) As Long

'Sub Sub0()
'End Sub

'Function Function0_1(ByVal arg As String) As String
Function Function0_1(ByVal arg As String) As String
WriteLogSimple "test1.bas:Function0_1 START"  'for DEBUG

WriteLogSimple "test1.bas:Function0_1 END 1"  'for DEBUG
	Exit Function

WriteLogSimple "test1.bas:Function0_1 END 2"  'for DEBUG
	Exit Function

WriteLogSimple "test1.bas:Function0_1 END"  'for DEBUG
End Function

Sub Sub0_1(ByVal arg As String) As String
WriteLogSimple "test1.bas:Sub0_1 START"  'for DEBUG

WriteLogSimple "test1.bas:Sub0_1 END 1"  'for DEBUG
	Exit Sub

WriteLogSimple "test1.bas:Sub0_1 END"  'for DEBUG
End Function

Function Function0_2(ByVal arg As String) As String 'hoge
WriteLogSimple "test1.bas:Function0_2 START"  'for DEBUG
WriteLogSimple "test1.bas:Function0_2 END"  'for DEBUG
End Function

Function Function0_3(ByVal arg As String _
) As String 'hoge
WriteLogSimple "test1.bas:Function0_3 START"  'for DEBUG
WriteLogSimple "test1.bas:Function0_3 END"  'for DEBUG
End Function

Function Function0_4(ByVal arg As String _
) As String
WriteLogSimple "test1.bas:Function0_4 START"  'for DEBUG
WriteLogSimple "test1.bas:Function0_4 END"  'for DEBUG
End Function

Function Function0_5(ByVal arg As String _
) As String()
WriteLogSimple "test1.bas:Function0_5 START"  'for DEBUG
WriteLogSimple "test1.bas:Function0_5 END"  'for DEBUG
End Function

Function Function0_6(ByVal arg As String _
) As String()	'hoge
WriteLogSimple "test1.bas:Function0_6 START"  'for DEBUG
WriteLogSimple "test1.bas:Function0_6 END"  'for DEBUG
End Function

Function Function0_7(ByVal arg As String _
) As _
String()	'hoge
WriteLogSimple "test1.bas:Function0_7 START"  'for DEBUG
WriteLogSimple "test1.bas:Function0_7 END"  'for DEBUG
End Function

Function Function0_8(ByVal arg As String _
) As _
String()
WriteLogSimple "test1.bas:Function0_8 START"  'for DEBUG
WriteLogSimple "test1.bas:Function0_8 END"  'for DEBUG
End Function

Function Function0_9(ByVal arg As String _
)
WriteLogSimple "test1.bas:Function0_9 START"  'for DEBUG
WriteLogSimple "test1.bas:Function0_9 END"  'for DEBUG
End Function

Function Function0_10()
WriteLogSimple "test1.bas:Function0_10 START"  'for DEBUG
WriteLogSimple "test1.bas:Function0_10 END"  'for DEBUG
End Function

Function Function0_10( _
)
WriteLogSimple "test1.bas:Function0_10 START"  'for DEBUG
WriteLogSimple "test1.bas:Function0_10 END"  'for DEBUG
End Function

Function Function0_11(ByVal arg As String _
) _
As _
String()
WriteLogSimple "test1.bas:Function0_11 START"  'for DEBUG
WriteLogSimple "test1.bas:Function0_11 END"  'for DEBUG
End Function

Sub Sub0_1(ByVal arg As String)
WriteLogSimple "test1.bas:Sub0_1 START"  'for DEBUG
WriteLogSimple "test1.bas:Sub0_1 END"  'for DEBUG
End Sub

Sub Sub0_2(ByVal arg As String)	'hoge
WriteLogSimple "test1.bas:Sub0_2 START"  'for DEBUG
WriteLogSimple "test1.bas:Sub0_2 END"  'for DEBUG
End Sub

Sub Sub0_3(ByVal arg As String _
)
WriteLogSimple "test1.bas:Sub0_3 START"  'for DEBUG
WriteLogSimple "test1.bas:Sub0_3 END"  'for DEBUG
End Sub

Sub Sub0_4(ByVal arg As String _
)	'hoge
WriteLogSimple "test1.bas:Sub0_4 START"  'for DEBUG
WriteLogSimple "test1.bas:Sub0_4 END"  'for DEBUG
End Sub

Sub Sub0(ByVal arg As String. _
         ByRef arg2 As Object)  'hoge
WriteLogSimple "test1.bas:Sub0 START"  'for DEBUG

	Dim a As Long
	If True Then
WriteLogSimple "test1.bas:Sub0 END 1"  'for DEBUG
		Exit Sub
	End If

WriteLogSimple "test1.bas:Sub0 END"  'for DEBUG
End Sub

Sub Sub1()
WriteLogSimple "test1.bas:Sub1 START"  'for DEBUG
WriteLogSimple "test1.bas:Sub1 END"  'for DEBUG
End Sub

Private Sub Sub2()
WriteLogSimple "test1.bas:Sub2 START"  'for DEBUG
WriteLogSimple "test1.bas:Sub2 END"  'for DEBUG
End Sub

Public Sub Sub3()
WriteLogSimple "test1.bas:Sub3 START"  'for DEBUG
WriteLogSimple "test1.bas:Sub3 END"  'for DEBUG
End Sub

Public Sub Sub4(ByVal arg As String)
WriteLogSimple "test1.bas:Sub4 START"  'for DEBUG
WriteLogSimple "test1.bas:Sub4 END"  'for DEBUG
End Sub

Public Sub Sub5(arg As String)
WriteLogSimple "test1.bas:Sub5 START"  'for DEBUG
WriteLogSimple "test1.bas:Sub5 END"  'for DEBUG
End Sub

Public Sub Sub6(arg)
WriteLogSimple "test1.bas:Sub6 START"  'for DEBUG
WriteLogSimple "test1.bas:Sub6 END"  'for DEBUG
End Sub



