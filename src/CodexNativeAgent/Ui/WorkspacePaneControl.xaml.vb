Imports System.Windows
Imports System.Windows.Controls

Namespace CodexNativeAgent.Ui
    Public Partial Class WorkspacePaneControl
        Inherits UserControl

        Private Const TurnComposerInputMinHeight As Double = 56.0R

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub WorkspacePaneControl_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
            UpdateTurnComposerResizeBounds()
        End Sub

        Private Sub WorkspacePaneControl_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
            UpdateTurnComposerResizeBounds()
        End Sub

        Private Sub TurnComposerSupplementalSection_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles InlineApprovalCard.SizeChanged, TurnComposerControlsHost.SizeChanged
            UpdateTurnComposerResizeBounds()
        End Sub

        Private Sub InlineApprovalCard_IsVisibleChanged(sender As Object, e As DependencyPropertyChangedEventArgs) Handles InlineApprovalCard.IsVisibleChanged
            UpdateTurnComposerResizeBounds()
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

            If TurnComposerInputHost Is Nothing OrElse TxtTurnInput Is Nothing Then
                Return
            End If

            Dim inputMinHeight = Math.Max(TurnComposerInputMinHeight, TurnComposerInputHost.MinHeight)
            Dim reservedHeight = MeasureOuterHeight(InlineApprovalCard) + MeasureOuterHeight(TurnComposerControlsHost)
            Dim maxInputHeight = Math.Max(inputMinHeight, Math.Floor(maxComposerHeight - reservedHeight))

            TurnComposerInputHost.MaxHeight = maxInputHeight
            TxtTurnInput.MaxHeight = maxInputHeight
        End Sub

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
