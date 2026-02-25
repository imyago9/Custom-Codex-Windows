Imports System.Windows
Imports System.Windows.Controls

Namespace CodexNativeAgent.Ui
    Public Partial Class WorkspacePaneControl
        Inherits UserControl

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub WorkspacePaneControl_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
            UpdateTurnComposerResizeBounds()
        End Sub

        Private Sub WorkspacePaneControl_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
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
        End Sub
    End Class
End Namespace
