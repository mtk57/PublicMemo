Public Class Form1
	Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		Dim td As New ToolTipOnDisabledControl()
		td.SetToolTip(Me.button2, Me.ToolTip1, "This tooltip is Disabled for Button2")

	End Sub



	Private Sub checkBox1_CheckedChanged(sender As Object, e As EventArgs) Handles checkBox1.CheckedChanged
		button1.Enabled = Not button1.Enabled

		If Not button1.Enabled Then
			Dim td As New ToolTipOnDisabledControl()
			'td.SetToolTip(this.button1, "1");
			'test
			td.SetToolTip(Me.button1, Me.ToolTip1, "This tooltip is Disabled for Button1")
		Else
			'test
			Me.ToolTip1.SetToolTip(Me.button1, "This tooltip is Enabled for Button1")
		End If
	End Sub

	''' <summary>
	''' // Reference example
	'''  var td = new ToolTipOnDisabledControl();
	'''  this.checkEdit3.Enabled = false;
	'''  td.SetTooltip(this.checkEdit3, "tooltip for disabled");
	''' </summary>
	Public Class ToolTipOnDisabledControl
#Region "Fields and Properties"

		Private enabledParentControl As Control

		Private isShown As Boolean

		Public Property TargetControl() As Control

		Public Property TooltipText() As String
		Public ReadOnly Property ToolTip() As ToolTip
#End Region

#Region "Public Methods"
		Public Sub New()
			Me.ToolTip = New ToolTip()

			Dim timer As New Timer()
			timer.Interval = 100 ' 100ミリ秒ごとに更新
			AddHandler timer.Tick, AddressOf Timer_Tick
			timer.Start()
		End Sub

		Private Sub Timer_Tick(ByVal sender As Object, ByVal e As EventArgs)
			' マウスカーソルの座標を取得して表示する
			'Dim cursorPosition As Point = Cursor.Position
			'Me.Label1.Text = $"X: {cursorPosition.X}, Y: {cursorPosition.Y}"

			' マウスカーソルのスクリーン座標を取得
			Dim screenPoint As Point = Cursor.Position

			' コントロールに対して座標変換を行う
			Dim clientPoint As Point = Me.enabledParentControl.PointToClient(screenPoint)

			If clientPoint.X >= Me.TargetControl.Left AndAlso
			   clientPoint.X <= Me.TargetControl.Right AndAlso
			   clientPoint.Y >= Me.TargetControl.Top AndAlso
			   clientPoint.Y <= Me.TargetControl.Bottom Then
				If Not Me.isShown Then
					Me.ToolTip.Show(
						Me.TooltipText,
						Me.TargetControl,
						Me.TargetControl.Width / 2, Me.TargetControl.Height / 2,
						Me.ToolTip.AutoPopDelay)
					Me.isShown = True
				End If
			Else
				'else if(this.isShown)
				Me.ToolTip.Hide(Me.TargetControl)
				Me.isShown = False
			End If
		End Sub

		'public void SetToolTip(Control targetControl, string tooltipText = null)
		Public Sub SetToolTip(targetControl As Control, tip As ToolTip, Optional tooltipText As String = Nothing)
			'test
			Me.TargetControl = targetControl
			If String.IsNullOrEmpty(tooltipText) Then
				Me.TooltipText = Me.ToolTip.GetToolTip(targetControl)
			Else
				Me.TooltipText = tooltipText

				' test
				tip.SetToolTip(targetControl, "")
			End If

			If targetControl.Enabled Then
				Me.enabledParentControl = Nothing
				Me.isShown = False
				Me.ToolTip.SetToolTip(Me.TargetControl, Me.TooltipText)
				Return
			End If

			Me.enabledParentControl = targetControl.Parent
			While Not Me.enabledParentControl.Enabled AndAlso Me.enabledParentControl.Parent IsNot Nothing
				Me.enabledParentControl = Me.enabledParentControl.Parent
			End While

			If Not Me.enabledParentControl.Enabled Then
				Throw New Exception("Failed to set tool tip because failed to find an enabled parent control.")
			End If

			AddHandler Me.enabledParentControl.MouseMove, AddressOf Me.EnabledParentControl_MouseMove
			AddHandler Me.TargetControl.EnabledChanged, AddressOf Me.TargetControl_EnabledChanged
		End Sub

		Public Sub Reset()
			If Me.TargetControl IsNot Nothing Then
				Me.ToolTip.Hide(Me.TargetControl)
				RemoveHandler Me.TargetControl.EnabledChanged, AddressOf Me.TargetControl_EnabledChanged
				Me.TargetControl = Nothing
			End If

			If Me.enabledParentControl IsNot Nothing Then
				RemoveHandler Me.enabledParentControl.MouseMove, AddressOf Me.EnabledParentControl_MouseMove
				Me.enabledParentControl = Nothing
			End If

			Me.isShown = False
		End Sub
#End Region

#Region "Private Methods"
		Private Sub EnabledParentControl_MouseMove(sender As Object, e As MouseEventArgs)
			'If e.Location.X >= Me.TargetControl.Left AndAlso
			'   e.Location.X <= Me.TargetControl.Right AndAlso
			'   e.Location.Y >= Me.TargetControl.Top AndAlso
			'   e.Location.Y <= Me.TargetControl.Bottom Then
			'	If Not Me.isShown Then
			'		Me.ToolTip.Show(
			'			Me.TooltipText,
			'			Me.TargetControl,
			'			Me.TargetControl.Width / 2, Me.TargetControl.Height / 2,
			'			Me.ToolTip.AutoPopDelay)
			'		Me.isShown = True
			'	End If
			'Else
			'	'else if(this.isShown)
			'	Me.ToolTip.Hide(Me.TargetControl)
			'	Me.isShown = False
			'End If
		End Sub

		Private Sub TargetControl_EnabledChanged(sender As Object, e As EventArgs)
			If TargetControl.Enabled Then
				RemoveHandler TargetControl.EnabledChanged, AddressOf TargetControl_EnabledChanged
				RemoveHandler enabledParentControl.MouseMove, AddressOf EnabledParentControl_MouseMove
			End If
		End Sub
#End Region
	End Class
End Class
