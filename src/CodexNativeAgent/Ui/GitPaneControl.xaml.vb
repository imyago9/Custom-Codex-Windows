Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives

Namespace CodexNativeAgent.Ui
    Public Partial Class GitPaneControl
        Private _isClampingVerticalSplitter As Boolean

        Public Sub New()
            InitializeComponent()

            AddHandler Loaded, AddressOf OnLoaded
        End Sub

        Private Sub OnLoaded(sender As Object, e As RoutedEventArgs)
            RemoveHandler Loaded, AddressOf OnLoaded

            If GitChangesDiffRowSplitter IsNot Nothing Then
                AddHandler GitChangesDiffRowSplitter.DragDelta, AddressOf OnGitChangesDiffRowSplitterDragDelta
                AddHandler GitChangesDiffRowSplitter.DragCompleted, AddressOf OnGitChangesDiffRowSplitterDragCompleted
            End If
            If GitHistoryPreviewRowSplitter IsNot Nothing Then
                AddHandler GitHistoryPreviewRowSplitter.DragDelta, AddressOf OnGitHistoryPreviewRowSplitterDragDelta
                AddHandler GitHistoryPreviewRowSplitter.DragCompleted, AddressOf OnGitHistoryPreviewRowSplitterDragCompleted
            End If
            If GitBranchesPreviewRowSplitter IsNot Nothing Then
                AddHandler GitBranchesPreviewRowSplitter.DragDelta, AddressOf OnGitBranchesPreviewRowSplitterDragDelta
                AddHandler GitBranchesPreviewRowSplitter.DragCompleted, AddressOf OnGitBranchesPreviewRowSplitterDragCompleted
            End If

            If GitTabChangesView IsNot Nothing Then
                AddHandler GitTabChangesView.SizeChanged, AddressOf OnGitTabChangesViewSizeChanged
            End If
            If GitTabHistoryView IsNot Nothing Then
                AddHandler GitTabHistoryView.SizeChanged, AddressOf OnGitTabHistoryViewSizeChanged
            End If
            If GitTabBranchesView IsNot Nothing Then
                AddHandler GitTabBranchesView.SizeChanged, AddressOf OnGitTabBranchesViewSizeChanged
            End If

            ClampChangesTabSplitter()
            ClampHistoryTabSplitter()
            ClampBranchesTabSplitter()
        End Sub

        Private Sub OnGitChangesDiffRowSplitterDragDelta(sender As Object, e As DragDeltaEventArgs)
            ApplyChangesTabSplitterDrag(e.VerticalChange)
        End Sub

        Private Sub OnGitChangesDiffRowSplitterDragCompleted(sender As Object, e As DragCompletedEventArgs)
            ClampChangesTabSplitter()
        End Sub

        Private Sub OnGitHistoryPreviewRowSplitterDragDelta(sender As Object, e As DragDeltaEventArgs)
            ApplyHistoryTabSplitterDrag(e.VerticalChange)
        End Sub

        Private Sub OnGitHistoryPreviewRowSplitterDragCompleted(sender As Object, e As DragCompletedEventArgs)
            ClampHistoryTabSplitter()
        End Sub

        Private Sub OnGitBranchesPreviewRowSplitterDragDelta(sender As Object, e As DragDeltaEventArgs)
            ApplyBranchesTabSplitterDrag(e.VerticalChange)
        End Sub

        Private Sub OnGitBranchesPreviewRowSplitterDragCompleted(sender As Object, e As DragCompletedEventArgs)
            ClampBranchesTabSplitter()
        End Sub

        Private Sub OnGitTabChangesViewSizeChanged(sender As Object, e As SizeChangedEventArgs)
            ClampChangesTabSplitter()
        End Sub

        Private Sub OnGitTabHistoryViewSizeChanged(sender As Object, e As SizeChangedEventArgs)
            ClampHistoryTabSplitter()
        End Sub

        Private Sub OnGitTabBranchesViewSizeChanged(sender As Object, e As SizeChangedEventArgs)
            ClampBranchesTabSplitter()
        End Sub

        Private Sub ClampChangesTabSplitter()
            ClampVerticalSplit(GitTabChangesView, GitChangesWorkingTreeRow, GitChangesDividerRow, GitChangesDiffRow)
        End Sub

        Private Sub ClampHistoryTabSplitter()
            ClampVerticalSplit(GitTabHistoryView, GitHistoryListRow, GitHistoryDividerRow, GitHistoryPreviewRow)
        End Sub

        Private Sub ClampBranchesTabSplitter()
            ClampVerticalSplit(GitTabBranchesView, GitBranchesListRow, GitBranchesDividerRow, GitBranchesPreviewRow)
        End Sub

        Private Sub ApplyChangesTabSplitterDrag(verticalChange As Double)
            ApplyVerticalSplitDrag(GitTabChangesView, GitChangesWorkingTreeRow, GitChangesDividerRow, GitChangesDiffRow, verticalChange)
        End Sub

        Private Sub ApplyHistoryTabSplitterDrag(verticalChange As Double)
            ApplyVerticalSplitDrag(GitTabHistoryView, GitHistoryListRow, GitHistoryDividerRow, GitHistoryPreviewRow, verticalChange)
        End Sub

        Private Sub ApplyBranchesTabSplitterDrag(verticalChange As Double)
            ApplyVerticalSplitDrag(GitTabBranchesView, GitBranchesListRow, GitBranchesDividerRow, GitBranchesPreviewRow, verticalChange)
        End Sub

        Private Sub ClampVerticalSplit(container As Grid,
                                       topRow As RowDefinition,
                                       dividerRow As RowDefinition,
                                       bottomRow As RowDefinition)
            If _isClampingVerticalSplitter Then
                Return
            End If

            If container Is Nothing OrElse topRow Is Nothing OrElse dividerRow Is Nothing OrElse bottomRow Is Nothing Then
                Return
            End If

            If container.Visibility <> Visibility.Visible OrElse container.ActualHeight <= 0 Then
                Return
            End If

            Dim availableResizableHeight = container.ActualHeight -
                                           container.RowDefinitions(0).ActualHeight -
                                           dividerRow.ActualHeight
            If availableResizableHeight <= 0 Then
                Return
            End If

            Dim minTop = Math.Max(0, topRow.MinHeight)
            Dim minBottom = Math.Max(0, bottomRow.MinHeight)
            Dim maxTop = Math.Max(0, availableResizableHeight - minBottom)
            Dim effectiveMinTop = Math.Min(minTop, maxTop)

            Dim currentTop = topRow.ActualHeight
            If currentTop <= 0 Then
                currentTop = effectiveMinTop
            End If

            Dim clampedTop = Math.Max(effectiveMinTop, Math.Min(currentTop, maxTop))

            _isClampingVerticalSplitter = True
            Try
                topRow.Height = New GridLength(clampedTop, GridUnitType.Pixel)
                bottomRow.Height = New GridLength(1, GridUnitType.Star)
            Finally
                _isClampingVerticalSplitter = False
            End Try
        End Sub

        Private Sub ApplyVerticalSplitDrag(container As Grid,
                                           topRow As RowDefinition,
                                           dividerRow As RowDefinition,
                                           bottomRow As RowDefinition,
                                           verticalChange As Double)
            If _isClampingVerticalSplitter Then
                Return
            End If

            If container Is Nothing OrElse topRow Is Nothing OrElse dividerRow Is Nothing OrElse bottomRow Is Nothing Then
                Return
            End If

            If container.Visibility <> Visibility.Visible OrElse container.ActualHeight <= 0 Then
                Return
            End If

            Dim availableResizableHeight = container.ActualHeight -
                                           container.RowDefinitions(0).ActualHeight -
                                           dividerRow.ActualHeight
            If availableResizableHeight <= 0 Then
                Return
            End If

            Dim minTop = Math.Max(0, topRow.MinHeight)
            Dim minBottom = Math.Max(0, bottomRow.MinHeight)
            Dim maxTop = Math.Max(0, availableResizableHeight - minBottom)
            Dim effectiveMinTop = Math.Min(minTop, maxTop)

            Dim currentTop = topRow.ActualHeight
            If currentTop <= 0 Then
                currentTop = effectiveMinTop
            End If

            Dim proposedTop = currentTop + verticalChange
            Dim clampedTop = Math.Max(effectiveMinTop, Math.Min(proposedTop, maxTop))
            Dim clampedBottom = Math.Max(minBottom, availableResizableHeight - clampedTop)

            _isClampingVerticalSplitter = True
            Try
                topRow.Height = New GridLength(clampedTop, GridUnitType.Pixel)
                bottomRow.Height = New GridLength(clampedBottom, GridUnitType.Pixel)
            Finally
                _isClampingVerticalSplitter = False
            End Try
        End Sub

    End Class
End Namespace
