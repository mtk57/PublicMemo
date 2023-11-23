Option Explicit On

Public Class Form1

	'UPGRADE Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
WriteLogSimple "Form1.vb:Form1_Load START"  'for DEBUG

WriteLogSimple "Form1.vb:Form1_Load END"  'for DEBUG
    End Sub

    Private Sub Form1_Load(sender As Object,e As EventArgs) Handles MyBase.Load
WriteLogSimple "Form1.vb:Form1_Load START"  'for DEBUG

		If a Then
WriteLogSimple "Form1.vb:Form1_Load END 1"  'for DEBUG
			Return 123
		End If

		Try
		Catch
WriteLogSimple "Form1.vb:Form1_Load END 2"  'for DEBUG
			Throw
		End Try


WriteLogSimple "Form1.vb:Form1_Load END 3"  'for DEBUG
		Return

WriteLogSimple "Form1.vb:Form1_Load END"  'for DEBUG
    End Sub
End Class


