<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    '�t�H�[�����R���|�[�l���g�̈ꗗ���N���[���A�b�v���邽�߂� dispose ���I�[�o�[���C�h���܂��B
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
WriteLogSimple "Dispose START"  'for DEBUG
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
WriteLogSimple "Dispose END"  'for DEBUG
    End Sub

    'Windows �t�H�[�� �f�U�C�i�[�ŕK�v�ł��B
    Private components As System.ComponentModel.IContainer

    '����: �ȉ��̃v���V�[�W���� Windows �t�H�[�� �f�U�C�i�[�ŕK�v�ł��B
    'Windows �t�H�[�� �f�U�C�i�[���g�p���ĕύX�ł��܂��B  
    '�R�[�h �G�f�B�^�[���g���ĕύX���Ȃ��ł��������B
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
WriteLogSimple "InitializeComponent START"  'for DEBUG
        Me.SuspendLayout()
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(13.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

WriteLogSimple "InitializeComponent END"  'for DEBUG
    End Sub

End Class

