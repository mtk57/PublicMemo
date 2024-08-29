Private Declare PtrSafe Function GetPrivateProfileString Lib _
    "kernel32" Alias "GetPrivateProfileStringA" ( _
) As Long

Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
) As Long

'Sub Sub0()
'End Sub

'Function Function0_1(ByVal arg As String) As String
Function Function0_1(ByVal arg As String) As String
WriteLogSimple "test1.bas" & vbTab & "Function0_1" & vbTab & "START"  'for DEBUG

WriteLogSimple "test1.bas" & vbTab & "Function0_1" & vbTab & "END_1"  'for DEBUG
	Exit Function

WriteLogSimple "test1.bas" & vbTab & "Function0_1" & vbTab & "END_2"  'for DEBUG
	Exit Function

WriteLogSimple "test1.bas" & vbTab & "Function0_1" & vbTab & "END"  'for DEBUG
End Function

Sub Sub0_1(ByVal arg As String) As String
WriteLogSimple "test1.bas" & vbTab & "Sub0_1" & vbTab & "START"  'for DEBUG

WriteLogSimple "test1.bas" & vbTab & "Sub0_1" & vbTab & "END_1"  'for DEBUG
	Exit Sub

WriteLogSimple "test1.bas" & vbTab & "Sub0_1" & vbTab & "END"  'for DEBUG
End Function

Function Function0_2(ByVal arg As String) As String 'hoge
WriteLogSimple "test1.bas" & vbTab & "Function0_2" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Function0_2" & vbTab & "END"  'for DEBUG
End Function

Function Function0_3(ByVal arg As String _
) As String 'hoge
WriteLogSimple "test1.bas" & vbTab & "Function0_3" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Function0_3" & vbTab & "END"  'for DEBUG
End Function

Function Function0_4(ByVal arg As String _
) As String
WriteLogSimple "test1.bas" & vbTab & "Function0_4" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Function0_4" & vbTab & "END"  'for DEBUG
End Function

Function Function0_5(ByVal arg As String _
) As String()
WriteLogSimple "test1.bas" & vbTab & "Function0_5" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Function0_5" & vbTab & "END"  'for DEBUG
End Function

Function Function0_6(ByVal arg As String _
) As String()	'hoge
WriteLogSimple "test1.bas" & vbTab & "Function0_6" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Function0_6" & vbTab & "END"  'for DEBUG
End Function

Function Function0_7(ByVal arg As String _
) As _
String()	'hoge
WriteLogSimple "test1.bas" & vbTab & "Function0_7" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Function0_7" & vbTab & "END"  'for DEBUG
End Function

Function Function0_8(ByVal arg As String _
) As _
String()
WriteLogSimple "test1.bas" & vbTab & "Function0_8" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Function0_8" & vbTab & "END"  'for DEBUG
End Function

Function Function0_9(ByVal arg As String _
)
WriteLogSimple "test1.bas" & vbTab & "Function0_9" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Function0_9" & vbTab & "END"  'for DEBUG
End Function

Function Function0_10()
WriteLogSimple "test1.bas" & vbTab & "Function0_10" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Function0_10" & vbTab & "END"  'for DEBUG
End Function

Function Function0_10( _
)
WriteLogSimple "test1.bas" & vbTab & "Function0_10" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Function0_10" & vbTab & "END"  'for DEBUG
End Function

Function Function0_11(ByVal arg As String _
) _
As _
String()
WriteLogSimple "test1.bas" & vbTab & "Function0_11" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Function0_11" & vbTab & "END"  'for DEBUG
End Function

Sub Sub0_1(ByVal arg As String)
WriteLogSimple "test1.bas" & vbTab & "Sub0_1" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Sub0_1" & vbTab & "END"  'for DEBUG
End Sub

Sub Sub0_2(ByVal arg As String)	'hoge
WriteLogSimple "test1.bas" & vbTab & "Sub0_2" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Sub0_2" & vbTab & "END"  'for DEBUG
End Sub

Sub Sub0_3(ByVal arg As String _
)
WriteLogSimple "test1.bas" & vbTab & "Sub0_3" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Sub0_3" & vbTab & "END"  'for DEBUG
End Sub

Sub Sub0_4(ByVal arg As String _
)	'hoge
WriteLogSimple "test1.bas" & vbTab & "Sub0_4" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Sub0_4" & vbTab & "END"  'for DEBUG
End Sub

Sub Sub0(ByVal arg As String. _
         ByRef arg2 As Object)  'hoge
WriteLogSimple "test1.bas" & vbTab & "Sub0" & vbTab & "START"  'for DEBUG

	Dim a As Long
	If True Then
WriteLogSimple "test1.bas" & vbTab & "Sub0" & vbTab & "END_1"  'for DEBUG
		Exit Sub
	End If

WriteLogSimple "test1.bas" & vbTab & "Sub0" & vbTab & "END"  'for DEBUG
End Sub

Sub Sub1()
WriteLogSimple "test1.bas" & vbTab & "Sub1" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Sub1" & vbTab & "END"  'for DEBUG
End Sub

Private Sub Sub2()
WriteLogSimple "test1.bas" & vbTab & "Sub2" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Sub2" & vbTab & "END"  'for DEBUG
End Sub

Public Sub Sub3()
WriteLogSimple "test1.bas" & vbTab & "Sub3" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Sub3" & vbTab & "END"  'for DEBUG
End Sub

Public Sub Sub4(ByVal arg As String)
WriteLogSimple "test1.bas" & vbTab & "Sub4" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Sub4" & vbTab & "END"  'for DEBUG
End Sub

Public Sub Sub5(arg As String)
WriteLogSimple "test1.bas" & vbTab & "Sub5" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Sub5" & vbTab & "END"  'for DEBUG
End Sub

Public Sub Sub6(arg)
WriteLogSimple "test1.bas" & vbTab & "Sub6" & vbTab & "START"  'for DEBUG
WriteLogSimple "test1.bas" & vbTab & "Sub6" & vbTab & "END"  'for DEBUG
End Sub



