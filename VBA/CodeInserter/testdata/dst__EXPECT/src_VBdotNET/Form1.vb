Option Explicit On

Public Class Form1

	'UPGRADE Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
WriteLogSimple("Form1.vb" & vbTab & "Form1_Load" & vbTab & "START")  'for DEBUG

WriteLogSimple("Form1.vb" & vbTab & "Form1_Load" & vbTab & "END")  'for DEBUG
    End Sub

    Private Sub Form1_Load(sender As Object,e As EventArgs) Handles MyBase.Load
WriteLogSimple("Form1.vb" & vbTab & "Form1_Load" & vbTab & "START")  'for DEBUG

		If a Then
WriteLogSimple("Form1.vb" & vbTab & "Form1_Load" & vbTab & "END_1")  'for DEBUG
			Return 123
		End If

		Try
		Catch
WriteLogSimple("Form1.vb" & vbTab & "Form1_Load" & vbTab & "END_2")  'for DEBUG
			Throw
		End Try


WriteLogSimple("Form1.vb" & vbTab & "Form1_Load" & vbTab & "END_3")  'for DEBUG
		Return

WriteLogSimple("Form1.vb" & vbTab & "Form1_Load" & vbTab & "END")  'for DEBUG
    End Sub
End Class


