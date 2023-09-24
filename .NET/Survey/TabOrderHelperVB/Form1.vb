Public Class Form1
    Private _helper As TabOrderHelper.TabOrderHelper = Nothing

    Private Sub Form1_Load(sender As Object, e As EventArgs)
        _helper = New TabOrderHelper.TabOrderHelper(Me)
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        Dim control = Me.ActiveControl

        If keyData = Keys.Tab Then
            ' TABキーが押されたときの処理

            Dim nextControl = _helper.GetNextControl(control)
            nextControl.Focus()
            Return True ' イベントを処理済みとしてマークする

        ElseIf keyData = (Keys.Shift Or Keys.Tab) Then
            ' SHIFT+TABキーが押されたときの処理

            Dim prevControl = _helper.GetNextControl(control, False)
            prevControl.Focus()
            Return True ' イベントを処理済みとしてマークする

        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

End Class