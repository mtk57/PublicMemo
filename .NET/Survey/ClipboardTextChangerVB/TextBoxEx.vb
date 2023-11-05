Public Class TextBoxEx

    'Private ReadOnly _clipBoardHelper As New Utility.ClipboardHelper(Me)
    'Private _clipBoardHelper As Utility.ClipboardHelper

    'Public Sub New()
    '    InitializeComponent()

    '    'AddHandler _clipBoardHelper.UpdateClipboard, AddressOf OnClipBoardUpdate

    '    'BackColor = Color.AliceBlue
    '    'Text = "__hoge__"
    'End Sub

    Private ReadOnly _clipBoardHelper As New Utility.ClipboardHelper(Me)

    Protected Overrides Sub OnCreateControl()
        AddHandler _clipBoardHelper.UpdateClipboard, AddressOf OnClipBoardUpdate

        BackColor = Color.AliceBlue
        Text = "__hoge__"

        MyBase.OnCreateControl()
    End Sub

    Private Sub OnClipBoardUpdate(ByVal sender As System.Object, ByVal e As EventArgs)
        Dim clipboardText As String = Clipboard.GetText()
        If MyBase.Text = clipboardText Then
            Clipboard.SetText(clipboardText.Replace("_", ""))
        End If
    End Sub

End Class
