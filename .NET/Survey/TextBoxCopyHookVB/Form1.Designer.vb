<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TextBoxEx1 = New WindowsApp1.TextBoxEx()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBoxEx2 = New WindowsApp1.TextBoxEx()
        Me.SuspendLayout()
        '
        'TextBoxEx1
        '
        Me.TextBoxEx1.Location = New System.Drawing.Point(99, 85)
        Me.TextBoxEx1.Name = "TextBoxEx1"
        Me.TextBoxEx1.Size = New System.Drawing.Size(249, 31)
        Me.TextBoxEx1.TabIndex = 0
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(114, 199)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(247, 31)
        Me.TextBox1.TabIndex = 1
        '
        'TextBoxEx2
        '
        Me.TextBoxEx2.Location = New System.Drawing.Point(404, 85)
        Me.TextBoxEx2.Name = "TextBoxEx2"
        Me.TextBoxEx2.Size = New System.Drawing.Size(249, 31)
        Me.TextBoxEx2.TabIndex = 2
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(13.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1069, 450)
        Me.Controls.Add(Me.TextBoxEx2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.TextBoxEx1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBoxEx1 As TextBoxEx
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBoxEx2 As TextBoxEx
End Class
