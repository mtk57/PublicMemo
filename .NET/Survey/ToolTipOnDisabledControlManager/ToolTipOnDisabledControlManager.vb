Namespace Utility
    Public NotInheritable Class ToolTipOnDisabledControlManager
        Private targetControls_ As System.Collections.Generic.List(Of TargetControl)
        Private toolTip_ As System.Windows.Forms.ToolTip
        Private timer_ As System.Windows.Forms.Timer

        Private Sub New()
            'Do nothing
        End Sub

        Public Sub New(ByRef toolTip As System.Windows.Forms.ToolTip)
            targetControls_ = New System.Collections.Generic.List(Of TargetControl)
            toolTip_ = toolTip
            timer_ = Nothing
        End Sub

        Public Sub Append(ByRef c As System.Windows.Forms.Control,
                          ByVal enabledText As String, ByVal disabledText As String)
            toolTip_.SetToolTip(c, "")

            targetControls_.Add(New TargetControl(c, toolTip_, enabledText, disabledText))
        End Sub

        Public Sub Update(ByRef c As System.Windows.Forms.Control,
                          ByVal enabledText As String, ByVal disabledText As String)

            For Each target In targetControls_
                If target.TargetControl.Name = c.Name Then
                    target.UpdateToolTipText(enabledText, disabledText)
                    Return
                End If
            Next

            Throw New System.Exception($"The specified control is not managed. Name={c.Name}")
        End Sub

        Public Sub Start()
            For Each target In targetControls_
                target.AddEnabledChangedEvent
            Next

            If timer_ IsNot Nothing Then
                timer_.Stop()
                RemoveHandler timer_.Tick, AddressOf Timer_Tick
                timer_ = Nothing
            End If

            timer_ = New System.Windows.Forms.Timer
            timer_.Interval = 100
            AddHandler timer_.Tick, AddressOf Timer_Tick
            timer_.Start()
        End Sub

        Public Sub [Stop]()
            For Each target In targetControls_
                target.RemoveEnabledChangedEvent()
            Next

            If timer_ IsNot Nothing Then
                timer_.Stop()
                RemoveHandler timer_.Tick, AddressOf Timer_Tick
                timer_ = Nothing
            End If
        End Sub

        Private Sub Timer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
            For Each target In targetControls_
                Dim c As System.Windows.Forms.Control = target.TargetControl
                If c.Enabled OrElse Not c.Visible Then
                    Continue For
                End If

                Dim scrPos As System.Drawing.Point = System.Windows.Forms.Cursor.Position

                Dim cliPos As System.Drawing.Point = c.PointToClient(scrPos)

                Dim left As Integer = 0
                Dim right As Integer = c.Width
                Dim top As Integer = 0
                Dim bottom As Integer = c.Height

                If cliPos.X >= left AndAlso
                   cliPos.X <= right AndAlso
                   cliPos.Y >= top AndAlso
                   cliPos.Y <= bottom Then

                    If Not target.IsShowDisabledToolTip Then
                        target.ShowDisabledToolTip(c)
                    End If
                Else
                    target.HideDisabledToolTip(c)
                End If
            Next
        End Sub

        Private NotInheritable Class TargetControl
            Private targetControl_ As System.Windows.Forms.Control
            Private enabledText_ As String
            Private disabledText_ As String
            Private isShowDisabledToolTip_ As Boolean
            Private toolTip_ As System.Windows.Forms.ToolTip

            Private Sub New()
                'do nothing
            End Sub

            Public Sub New(ByRef c As System.Windows.Forms.Control,
                           ByRef toolTip As System.Windows.Forms.ToolTip,
                           ByVal enabledText As String, ByVal disabledText As String)
                toolTip_ = New System.Windows.Forms.ToolTip()

                toolTip_.AutomaticDelay = toolTip.AutomaticDelay
                toolTip_.AutoPopDelay = toolTip.AutoPopDelay
                toolTip_.BackColor = toolTip.BackColor
                toolTip_.ForeColor = toolTip.ForeColor
                toolTip_.InitialDelay = toolTip.InitialDelay
                toolTip_.IsBalloon = toolTip.IsBalloon
                toolTip_.OwnerDraw = toolTip.OwnerDraw
                toolTip_.ReshowDelay = toolTip.ReshowDelay
                toolTip_.ShowAlways = toolTip.ShowAlways
                toolTip_.StripAmpersands = toolTip.StripAmpersands
                toolTip_.ToolTipTitle = toolTip.ToolTipTitle
                toolTip_.UseAnimation = toolTip.UseAnimation
                toolTip_.UseFading = toolTip.UseFading

                toolTip_.SetToolTip(c, "")

                isShowDisabledToolTip_ = False
                targetControl_ = c

                enabledText_ = enabledText
                disabledText_ = disabledText

                If targetControl_.Enabled Then
                    toolTip_.SetToolTip(targetControl_, enabledText_)
                Else
                    toolTip_.SetToolTip(targetControl_, disabledText_)
                End If
            End Sub

            Private Sub TargetControl_EnabledChanged(sender As System.Object, e As System.EventArgs)
                Dim c As System.Windows.Forms.Control = DirectCast(sender, System.Windows.Forms.Control)

                If c.Enabled Then
                    HideDisabledToolTip(c)
                End If
            End Sub

            Public Sub AddEnabledChangedEvent()
                RemoveHandler targetControl_.EnabledChanged, AddressOf TargetControl_EnabledChanged
                AddHandler targetControl_.EnabledChanged, AddressOf TargetControl_EnabledChanged
            End Sub

            Public Sub RemoveEnabledChangedEvent()
                RemoveHandler targetControl_.EnabledChanged, AddressOf TargetControl_EnabledChanged
            End Sub

            Public Sub ShowDisabledToolTip(ByRef c As System.Windows.Forms.Control)
                Dim pos As System.Drawing.Point = c.PointToClient(System.Windows.Forms.Cursor.Position)
                Const DISP_OFFSET As Integer = 18
                pos.Y = pos.Y + DISP_OFFSET

                toolTip_.Show(disabledText_, c, pos, toolTip_.AutoPopDelay)
                isShowDisabledToolTip_ = True
            End Sub

            Public Sub HideDisabledToolTip(ByRef c As System.Windows.Forms.Control)
                toolTip_.Hide(c)
                toolTip_.SetToolTip(c, enabledText_)
                isShowDisabledToolTip_ = False
            End Sub

            Public Sub UpdateToolTipText(ByVal enabledText As String, ByVal disabledTest As String)
                enabledText_ = enabledText
                disabledText_ = disabledTest

                If targetControl_.Enabled Then
                    toolTip_.SetToolTip(targetControl_, enabledText_)
                Else
                    toolTip_.SetToolTip(targetControl_, disabledText_)
                End If
            End Sub

            Public ReadOnly Property TargetControl As System.Windows.Forms.Control
                Get
                    Return targetControl_
                End Get
            End Property

            Public ReadOnly Property IsShowDisabledToolTip As Boolean
                Get
                    Return isShowDisabledToolTip_
                End Get
            End Property

        End Class
    End Class
End Namespace