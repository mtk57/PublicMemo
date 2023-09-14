Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim timer As New Timer()
        timer.Interval = 100 ' 100ミリ秒ごとに更新
        AddHandler timer.Tick, AddressOf Timer_Tick
        timer.Start()
    End Sub

    Private Sub Timer_Tick(ByVal sender As Object, ByVal e As EventArgs)
        ' マウスカーソルの座標を取得して表示する
        Dim cursorPosition As Point = Cursor.Position
        Me.Label1.Text = $"X: {cursorPosition.X}, Y: {cursorPosition.Y}"
    End Sub
End Class
