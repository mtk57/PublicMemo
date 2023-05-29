Imports System.Reflection


'文字列リソースを動的に取得する
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            MessageBox.Show(
                My.Resources.ResourceManager.GetString("String1") = GetStringResource_Dynamic(1)
            )
        Catch ex As Exception
            MessageBox.Show(ex.Message & ", " & ex.StackTrace)
        End Try
    End Sub

    Private Function GetStringResource_Dynamic(ByVal num As Integer) As String
        Dim t As Type
        Dim pi As PropertyInfo = Nothing

        t = Type.GetType(Me.GetType().Namespace & ".My.Resources.Resources")
        If t Is Nothing Then
            Return String.Empty
        End If

        For Each item In t.GetProperties(BindingFlags.NonPublic Or BindingFlags.Static)
            If item.Name = "ResourceManager" Then
                pi = item
                Exit For
            End If
        Next

        If pi Is Nothing Then
            Return String.Empty
        End If

        Return pi.GetValue(Nothing).GetString("String" & num)
    End Function
End Class
