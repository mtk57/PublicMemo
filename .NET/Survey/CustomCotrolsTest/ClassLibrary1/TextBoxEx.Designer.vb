Imports System.ComponentModel
Imports System.Drawing

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class TextBoxEx
    Inherits System.Windows.Forms.TextBox

    Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        BackColor = Color.Yellow
    End Sub

    <DefaultValue(GetType(Color), "Yellow")>
    Public Overrides Property BackColor As Color
        Get
            Return MyBase.BackColor
        End Get
        Set(value As Color)
            MyBase.BackColor = value
        End Set
    End Property

    'Control は、コンポーネント一覧に後処理を実行するために、dispose をオーバーライドします。
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

    'コントロール デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    ' メモ: 以下のプロシージャはコンポーネント デザイナーで必要です。
    ' コンポーネント デザイナーを使って変更できます。
    ' コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
    End Sub

End Class

