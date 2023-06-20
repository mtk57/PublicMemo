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
        Me.LabelEx1 = New ClassLibrary1.LabelEx()
        Me.TextBoxEx1 = New ClassLibrary1.TextBoxEx()
        Me.SuspendLayout()
        '
        'LabelEx1
        '
        Me.LabelEx1.AutoSize = True
        Me.LabelEx1.Location = New System.Drawing.Point(165, 76)
        Me.LabelEx1.Name = "LabelEx1"
        Me.LabelEx1.Size = New System.Drawing.Size(100, 24)
        Me.LabelEx1.TabIndex = 0
        Me.LabelEx1.Text = "LabelEx1"
        '
        'TextBoxEx1
        '
        Me.TextBoxEx1.Location = New System.Drawing.Point(182, 198)
        Me.TextBoxEx1.Name = "TextBoxEx1"
        Me.TextBoxEx1.Size = New System.Drawing.Size(100, 31)
        Me.TextBoxEx1.TabIndex = 1
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(13.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.TextBoxEx1)
        Me.Controls.Add(Me.LabelEx1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LabelEx1 As ClassLibrary1.LabelEx
    Friend WithEvents TextBoxEx1 As ClassLibrary1.TextBoxEx
End Class
