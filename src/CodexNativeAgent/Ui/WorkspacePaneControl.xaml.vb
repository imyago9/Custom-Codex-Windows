Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Threading
Imports CodexNativeAgent.Ui.ViewModels

Namespace CodexNativeAgent.Ui
    Public Partial Class WorkspacePaneControl
        Inherits UserControl

        Private Const TurnComposerInputMinHeight As Double = 56.0R
        Private Const TurnComposerInputBottomSafetyGutter As Double = 14.0R
        Private _composerRowHeightStoredForApproval As Boolean
        Private _composerRowHeightBeforeApproval As GridLength
        Private _deferredTurnComposerResizeQueued As Boolean

        Public Sub New()
            InitializeComponent()
            _composerRowHeightBeforeApproval = New GridLength(136.0R, GridUnitType.Pixel)
        End Sub

        Private Sub WorkspacePaneControl_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
            UpdateTurnComposerResizeBounds()
            QueueDeferredTurnComposerResize()
        End Sub

        Private Sub WorkspacePaneControl_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
            UpdateTurnComposerResizeBounds()
        End Sub

        Private Sub TxtTurnInput_TextChanged(sender As Object, e As TextChangedEventArgs) Handles TxtTurnInput.TextChanged
            UpdateTurnComposerResizeBounds()
            QueueDeferredTurnComposerResize()
        End Sub

        Private Sub TurnComposerSupplementalSection_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles InlineApprovalCard.SizeChanged, TurnComposerControlsHost.SizeChanged
            UpdateTurnComposerResizeBounds()
        End Sub

        Private Sub InlineApprovalCard_IsVisibleChanged(sender As Object, e As DependencyPropertyChangedEventArgs) Handles InlineApprovalCard.IsVisibleChanged
            UpdateTurnComposerResizeBounds()
            If InlineApprovalCard IsNot Nothing AndAlso InlineApprovalCard.Visibility = Visibility.Visible Then
                Keyboard.Focus(InlineApprovalCard)
            End If
        End Sub

        Private Sub QueueDeferredTurnComposerResize()
            If _deferredTurnComposerResizeQueued Then
                Return
            End If

            _deferredTurnComposerResizeQueued = True
            Dispatcher.BeginInvoke(
                DispatcherPriority.Render,
                New Action(
                    Sub()
                        _deferredTurnComposerResizeQueued = False
                        UpdateTurnComposerResizeBounds()
                    End Sub))
        End Sub

        Private Sub WorkspacePaneControl_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Me.PreviewKeyDown
            If e Is Nothing Then
                Return
            End If

            If InlineApprovalCard Is Nothing OrElse InlineApprovalCard.Visibility <> Visibility.Visible Then
                Return
            End If

            Dim mainVm = TryCast(DataContext, MainWindowViewModel)
            Dim approvalPanel = mainVm?.ApprovalPanel
            If approvalPanel Is Nothing Then
                Return
            End If

            Dim handled As Boolean
            Dim recognized As Boolean
            Select Case e.Key
                Case Key.Up
                    recognized = True
                    handled = approvalPanel.MoveSelection(-1)
                Case Key.Down
                    recognized = True
                    handled = approvalPanel.MoveSelection(1)
                Case Key.Enter, Key.Return
                    recognized = True
                    handled = approvalPanel.TryExecuteSelectedOption()
                Case Key.D1, Key.NumPad1
                    recognized = True
                    handled = approvalPanel.TryExecuteOption(1)
                Case Key.D2, Key.NumPad2
                    recognized = True
                    handled = approvalPanel.TryExecuteOption(2)
                Case Key.D3, Key.NumPad3
                    recognized = True
                    handled = approvalPanel.TryExecuteOption(3)
                Case Key.D4, Key.NumPad4
                    recognized = True
                    handled = approvalPanel.TryExecuteOption(4)
                Case Key.D5, Key.NumPad5
                    recognized = True
                    handled = approvalPanel.TryExecuteOption(5)
                Case Else
                    recognized = False
                    handled = False
            End Select

            If recognized OrElse handled Then
                e.Handled = True
            End If
        End Sub

        Private Sub UpdateTurnComposerResizeBounds()
            If WorkspaceLayoutRoot Is Nothing OrElse TurnComposerHostRow Is Nothing Then
                Return
            End If

            Dim workspaceHeight = WorkspaceLayoutRoot.ActualHeight
            If Double.IsNaN(workspaceHeight) OrElse Double.IsInfinity(workspaceHeight) OrElse workspaceHeight <= 0 Then
                Return
            End If

            Dim minComposerHeight = Math.Max(0.0R, TurnComposerHostRow.MinHeight)
            Dim maxComposerHeight = Math.Max(minComposerHeight, Math.Floor(workspaceHeight / 2.0R))
            TurnComposerHostRow.MaxHeight = maxComposerHeight

            Dim approvalVisible = InlineApprovalCard IsNot Nothing AndAlso
                                  InlineApprovalCard.Visibility = Visibility.Visible
            If approvalVisible Then
                If Not _composerRowHeightStoredForApproval Then
                    _composerRowHeightBeforeApproval = TurnComposerHostRow.Height
                    _composerRowHeightStoredForApproval = True
                End If
                ExpandComposerHeightForApproval(minComposerHeight, maxComposerHeight)
            ElseIf _composerRowHeightStoredForApproval Then
                TurnComposerHostRow.Height = _composerRowHeightBeforeApproval
                _composerRowHeightStoredForApproval = False
            End If

            If TurnComposerInputHost Is Nothing OrElse TxtTurnInput Is Nothing Then
                Return
            End If

            Dim inputMinHeight = Math.Max(TurnComposerInputMinHeight, TurnComposerInputHost.MinHeight)
            Dim reservedHeight = MeasureOuterHeight(InlineApprovalCard) + MeasureOuterHeight(TurnComposerControlsHost)
            Dim maxInputHeight = Math.Max(inputMinHeight, Math.Floor(maxComposerHeight - reservedHeight))

            TurnComposerInputHost.MaxHeight = maxInputHeight
            TxtTurnInput.MaxHeight = maxInputHeight

            If Not approvalVisible Then
                ExpandComposerHeightForInputContent(minComposerHeight, maxComposerHeight, inputMinHeight, maxInputHeight)
            End If
        End Sub

        Private Sub ExpandComposerHeightForApproval(minComposerHeight As Double,
                                                    maxComposerHeight As Double)
            If InlineApprovalCard Is Nothing OrElse TurnComposerHostRow Is Nothing Then
                Return
            End If

            Dim measureWidth = Math.Max(0.0R, InlineApprovalCard.ActualWidth)
            If measureWidth <= 1.0R AndAlso TurnComposerHostContainer IsNot Nothing Then
                measureWidth = Math.Max(0.0R,
                                        TurnComposerHostContainer.ActualWidth -
                                        InlineApprovalCard.Margin.Left -
                                        InlineApprovalCard.Margin.Right)
            End If

            If measureWidth <= 1.0R Then
                Return
            End If

            InlineApprovalCard.Measure(New Size(measureWidth, Double.PositiveInfinity))
            Dim desiredHeight = InlineApprovalCard.DesiredSize.Height +
                                InlineApprovalCard.Margin.Top +
                                InlineApprovalCard.Margin.Bottom
            If desiredHeight <= 0.0R Then
                Return
            End If

            Dim targetHeight = Math.Max(minComposerHeight,
                                        Math.Min(maxComposerHeight, Math.Ceiling(desiredHeight)))
            If TurnComposerHostRow.Height.IsAbsolute AndAlso
               Math.Abs(TurnComposerHostRow.Height.Value - targetHeight) < 0.5R Then
                Return
            End If

            TurnComposerHostRow.Height = New GridLength(targetHeight, GridUnitType.Pixel)
        End Sub

        Private Sub ExpandComposerHeightForInputContent(minComposerHeight As Double,
                                                        maxComposerHeight As Double,
                                                        inputMinHeight As Double,
                                                        inputMaxHeight As Double)
            If TurnComposerHostRow Is Nothing OrElse TurnComposerControlsHost Is Nothing OrElse TxtTurnInput Is Nothing Then
                Return
            End If

            Dim controlsHeight = MeasureOuterHeight(TurnComposerControlsHost)
            Dim inputAtExpansionLimit As Boolean
            Dim desiredInputHeight = ResolveTurnInputContentDesiredHeight(inputMinHeight,
                                                                          inputMaxHeight,
                                                                          inputAtExpansionLimit)
            Dim requiredComposerHeight = controlsHeight + desiredInputHeight
            requiredComposerHeight = Math.Max(minComposerHeight, Math.Min(maxComposerHeight, Math.Ceiling(requiredComposerHeight)))
            Dim currentHeight = ResolveCurrentComposerHeight(minComposerHeight)
            If Math.Abs(currentHeight - requiredComposerHeight) > 0.5R Then
                TurnComposerHostRow.Height = New GridLength(requiredComposerHeight, GridUnitType.Pixel)
            End If

            TxtTurnInput.VerticalScrollBarVisibility = If(inputAtExpansionLimit,
                                                          ScrollBarVisibility.Auto,
                                                          ScrollBarVisibility.Disabled)
        End Sub

        Private Function ResolveTurnInputContentDesiredHeight(inputMinHeight As Double,
                                                              inputMaxHeight As Double,
                                                              ByRef inputAtExpansionLimit As Boolean) As Double
            inputAtExpansionLimit = False
            If TxtTurnInput Is Nothing Then
                Return inputMinHeight
            End If

            Dim currentText = TxtTurnInput.Text
            If String.IsNullOrEmpty(currentText) Then
                Return inputMinHeight
            End If

            Dim extentHeight = TxtTurnInput.ExtentHeight
            If Double.IsNaN(extentHeight) OrElse Double.IsInfinity(extentHeight) OrElse extentHeight < 0.0R Then
                extentHeight = 0.0R
            End If

            Dim caretContentHeight As Double
            Dim caretRect = TxtTurnInput.GetRectFromCharacterIndex(Math.Max(0, currentText.Length), True)
            If Not Double.IsNaN(caretRect.Bottom) AndAlso Not Double.IsInfinity(caretRect.Bottom) Then
                caretContentHeight = caretRect.Bottom + Math.Ceiling(TxtTurnInput.FontSize * 0.35R)
            End If

            If Double.IsNaN(caretContentHeight) OrElse Double.IsInfinity(caretContentHeight) OrElse caretContentHeight < 0.0R Then
                caretContentHeight = 0.0R
            End If

            Dim chromeHeight = TxtTurnInput.Padding.Top +
                               TxtTurnInput.Padding.Bottom +
                               TxtTurnInput.Margin.Top +
                               TxtTurnInput.Margin.Bottom +
                               TurnComposerInputBottomSafetyGutter

            Dim desiredHeight = Math.Max(extentHeight, caretContentHeight)
            If desiredHeight <= 0.5R Then
                Dim measureWidth = Math.Max(0.0R, TxtTurnInput.ActualWidth)
                If measureWidth <= 1.0R AndAlso TurnComposerInputHost IsNot Nothing Then
                    measureWidth = Math.Max(0.0R,
                                            TurnComposerInputHost.ActualWidth -
                                            TxtTurnInput.Margin.Left -
                                            TxtTurnInput.Margin.Right)
                End If

                If measureWidth > 1.0R Then
                    TxtTurnInput.Measure(New Size(measureWidth, Double.PositiveInfinity))
                    Dim measuredFallbackHeight = TxtTurnInput.DesiredSize.Height -
                                                 TxtTurnInput.Padding.Top -
                                                 TxtTurnInput.Padding.Bottom -
                                                 TurnComposerInputBottomSafetyGutter
                    If Not Double.IsNaN(measuredFallbackHeight) AndAlso
                       Not Double.IsInfinity(measuredFallbackHeight) AndAlso
                       measuredFallbackHeight > 0.0R Then
                        desiredHeight = measuredFallbackHeight
                    End If
                End If
            End If

            Dim resolved = desiredHeight + chromeHeight
            If Double.IsNaN(resolved) OrElse Double.IsInfinity(resolved) OrElse resolved <= 0.0R Then
                resolved = inputMinHeight
            End If

            Dim normalizedResolved = Math.Max(inputMinHeight, Math.Ceiling(resolved))
            inputAtExpansionLimit = normalizedResolved > inputMaxHeight + 0.5R
            Return Math.Max(inputMinHeight, Math.Min(inputMaxHeight, normalizedResolved))
        End Function

        Private Function ResolveCurrentComposerHeight(minComposerHeight As Double) As Double
            If TurnComposerHostRow Is Nothing Then
                Return minComposerHeight
            End If

            If TurnComposerHostRow.Height.IsAbsolute Then
                Return TurnComposerHostRow.Height.Value
            End If

            Dim actual = TurnComposerHostRow.ActualHeight
            If Double.IsNaN(actual) OrElse Double.IsInfinity(actual) OrElse actual <= 0.0R Then
                Return minComposerHeight
            End If

            Return actual
        End Function

        Private Shared Function MeasureOuterHeight(element As FrameworkElement) As Double
            If element Is Nothing OrElse element.Visibility <> Visibility.Visible Then
                Return 0.0R
            End If

            Dim height = element.ActualHeight
            If Double.IsNaN(height) OrElse Double.IsInfinity(height) OrElse height <= 0 Then
                height = element.DesiredSize.Height
            End If

            If Double.IsNaN(height) OrElse Double.IsInfinity(height) OrElse height < 0 Then
                height = 0.0R
            End If

            height += element.Margin.Top + element.Margin.Bottom
            Return Math.Max(0.0R, height)
        End Function
    End Class
End Namespace
