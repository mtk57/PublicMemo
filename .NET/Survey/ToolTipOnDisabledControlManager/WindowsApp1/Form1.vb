
Public Class Form1

    Private mgr_ As ToolTipOnDisabledControlManager.Utility.ToolTipOnDisabledControlManager

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mgr_ = New ToolTipOnDisabledControlManager.Utility.ToolTipOnDisabledControlManager(ToolTip1)
        mgr_.Append(Button1, "Hoge", "Fuga")
        mgr_.Start()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Button1.Enabled = Not Button1.Enabled
    End Sub
End Class
