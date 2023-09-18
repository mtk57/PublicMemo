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
        Me.components = New System.ComponentModel.Container()
        Me.button1 = New System.Windows.Forms.Button()
        Me.checkBox1 = New System.Windows.Forms.CheckBox()
        Me.button2 = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.SuspendLayout()
        '
        'button1
        '
        Me.button1.Location = New System.Drawing.Point(193, 173)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(270, 91)
        Me.button1.TabIndex = 1
        Me.button1.Text = "button1"
        Me.ToolTip1.SetToolTip(Me.button1, "Hello button1")
        Me.button1.UseVisualStyleBackColor = True
        '
        'checkBox1
        '
        Me.checkBox1.AutoSize = True
        Me.checkBox1.Location = New System.Drawing.Point(203, 26)
        Me.checkBox1.Name = "checkBox1"
        Me.checkBox1.Size = New System.Drawing.Size(201, 28)
        Me.checkBox1.TabIndex = 2
        Me.checkBox1.Text = "Enabled Change"
        Me.checkBox1.UseVisualStyleBackColor = True
        '
        'button2
        '
        Me.button2.Enabled = False
        Me.button2.Location = New System.Drawing.Point(203, 334)
        Me.button2.Name = "button2"
        Me.button2.Size = New System.Drawing.Size(260, 92)
        Me.button2.TabIndex = 15
        Me.button2.Text = "button2"
        Me.ToolTip1.SetToolTip(Me.button2, "Hello button2")
        Me.button2.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(121, 100)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(457, 440)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "GroupBox1"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(13.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 710)
        Me.Controls.Add(Me.button2)
        Me.Controls.Add(Me.checkBox1)
        Me.Controls.Add(Me.button1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private WithEvents button1 As Button
    Friend WithEvents checkBox1 As CheckBox
    Private WithEvents button2 As Button
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents GroupBox1 As GroupBox
End Class
