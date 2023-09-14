Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles

Namespace ClassLibrary1
    Public Class CustomGroupBox
        Inherits GroupBox

        Public Sub New()
            SetStyle(ControlStyles.UserPaint, True)
            SetStyle(ControlStyles.DoubleBuffer, True)
            SetStyle(ControlStyles.ResizeRedraw, True)
        End Sub

        Private _DisabledColor As Color = Color.Red
        Private _BoxOffset As Point
        Private _CaptionOffset As Point

        <DefaultValue(GetType(System.Drawing.Color), "Red")>
        <Browsable(True)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
        Public Property DisabledColor As Color
            Get
                Return _DisabledColor
            End Get
            Set(value As Color)
                _DisabledColor = value
            End Set
        End Property

        <Browsable(True)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
        Public Property BoxOffset As Point
            Get
                Return _BoxOffset
            End Get
            Set(value As Point)
                _BoxOffset = value
            End Set
        End Property

        <Browsable(True)>
        <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
        Public Property CaptionOffset As Point
            Get
                Return _CaptionOffset
            End Get
            Set(value As Point)
                _CaptionOffset = value
            End Set
        End Property

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)

            If MyBase.Enabled Then Return

            Using brush As New SolidBrush(_DisabledColor)
                Using bkcol As New SolidBrush(BackColor)
                    If System.Windows.Forms.Application.RenderWithVisualStyles Then
                        e.Graphics.FillRectangle(bkcol, ClientRectangle)

                        GroupBoxRenderer.DrawGroupBox(
                            e.Graphics,
                            New Rectangle(
                                ClientRectangle.X + _BoxOffset.X, ClientRectangle.Y + _BoxOffset.Y,
                                ClientRectangle.Width - _BoxOffset.X, ClientRectangle.Height - _BoxOffset.Y),
                            GroupBoxState.Normal)

                        Using graphics As Graphics = graphics.FromImage(New Bitmap(1, 1))
                            Dim textSize As SizeF = graphics.MeasureString(Text, Font)

                            e.Graphics.FillRectangle(
                                bkcol,
                                New RectangleF(
                                    0 + _CaptionOffset.X, 0 + _CaptionOffset.Y,
                                    textSize.Width, textSize.Height))

                            e.Graphics.DrawString(
                                Text, Font, brush,
                                ClientRectangle.X + _CaptionOffset.X, ClientRectangle.Y + _CaptionOffset.Y)
                        End Using
                    End If
                End Using
            End Using
        End Sub

    End Class
End Namespace