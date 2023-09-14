<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.CustomGroupBox1 = New ClassLibrary1.ClassLibrary1.CustomGroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'CustomGroupBox1
        '
        Me.CustomGroupBox1.BoxOffset = New System.Drawing.Point(0, 6)
        Me.CustomGroupBox1.CaptionOffset = New System.Drawing.Point(6, 0)
        Me.CustomGroupBox1.Enabled = False
        Me.CustomGroupBox1.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.CustomGroupBox1.Location = New System.Drawing.Point(101, 66)
        Me.CustomGroupBox1.Name = "CustomGroupBox1"
        Me.CustomGroupBox1.Size = New System.Drawing.Size(497, 267)
        Me.CustomGroupBox1.TabIndex = 0
        Me.CustomGroupBox1.TabStop = False
        Me.CustomGroupBox1.Text = "CustomGroupBox1"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(101, 372)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(170, 35)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(13.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CustomGroupBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents CustomGroupBox1 As ClassLibrary1.ClassLibrary1.CustomGroupBox
    Friend WithEvents Button1 As Button
End Class
