Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Threading
Imports CodexNativeAgent.Ui.ViewModels.Threads

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private NotInheritable Class TranscriptTabSurfaceHandle
            Public Property ThreadId As String = String.Empty
            Public Property SurfaceListBox As ListBox
            Public Property TabButton As Button
            Public Property TabChipBorder As Border
            Public Property TabCloseButton As Button
            Public Property IsPrimarySurface As Boolean
        End Class

        Private NotInheritable Class TranscriptTabSurfaceRetireWorkItem
            Public Property ThreadId As String = String.Empty
            Public Property SurfaceListBox As ListBox
            Public Property IsPrimarySurface As Boolean
            Public Property QueuedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        End Class

        Private ReadOnly _transcriptTabSurfacesByThreadId As New Dictionary(Of String, TranscriptTabSurfaceHandle)(StringComparer.Ordinal)
        Private ReadOnly _transcriptInteractionHandlersAttached As New List(Of ListBox)()
        Private ReadOnly _transcriptTabSurfaceRetireQueue As New Queue(Of TranscriptTabSurfaceRetireWorkItem)()
        Private _transcriptSurfaceHostPanel As Panel
        Private _transcriptTabStripBorder As Border
        Private _transcriptTabStripScrollViewer As ScrollViewer
        Private _transcriptTabStripPanel As StackPanel
        Private _activeTranscriptSurfaceListBox As ListBox
        Private _activeTranscriptSurfaceThreadId As String = String.Empty
        Private _transcriptTabSurfaceRetireDrainScheduled As Boolean

        Private Function CurrentTranscriptListControl() As ListBox
            If _activeTranscriptSurfaceListBox IsNot Nothing Then
                Return _activeTranscriptSurfaceListBox
            End If

            If WorkspacePaneHost Is Nothing Then
                Return Nothing
            End If

            Return WorkspacePaneHost.LstTranscript
        End Function

        Private Sub AttachTranscriptInteractionHandlers(listBox As ListBox)
            If listBox Is Nothing Then
                Return
            End If

            For Each existing In _transcriptInteractionHandlersAttached
                If ReferenceEquals(existing, listBox) Then
                    Return
                End If
            Next

            listBox.AddHandler(ScrollViewer.ScrollChangedEvent,
                               New ScrollChangedEventHandler(AddressOf OnTranscriptScrollChanged))
            AddHandler listBox.PreviewMouseWheel,
                New MouseWheelEventHandler(AddressOf OnTranscriptPreviewMouseWheel)
            AddHandler listBox.PreviewMouseLeftButtonDown,
                New MouseButtonEventHandler(AddressOf OnTranscriptPreviewMouseLeftButtonDown)
            AddHandler listBox.PreviewKeyDown,
                New KeyEventHandler(AddressOf OnTranscriptPreviewKeyDown)
            listBox.AddHandler(Thumb.DragStartedEvent,
                               New DragStartedEventHandler(AddressOf OnTranscriptScrollThumbDragStarted),
                               True)
            listBox.AddHandler(Thumb.DragCompletedEvent,
                               New DragCompletedEventHandler(AddressOf OnTranscriptScrollThumbDragCompleted),
                               True)

            _transcriptInteractionHandlersAttached.Add(listBox)
        End Sub

        Private Sub EnsureTranscriptTabsUiInitialized()
            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.LstTranscript Is Nothing Then
                Return
            End If

            If _transcriptSurfaceHostPanel Is Nothing Then
                _transcriptSurfaceHostPanel = TryCast(WorkspacePaneHost.LstTranscript.Parent, Panel)
            End If

            If _activeTranscriptSurfaceListBox Is Nothing Then
                _activeTranscriptSurfaceListBox = WorkspacePaneHost.LstTranscript
            End If

            WorkspacePaneHost.LstTranscript.DataContext = _viewModel
            AttachTranscriptInteractionHandlers(WorkspacePaneHost.LstTranscript)

            If _transcriptTabStripPanel Is Nothing Then
                _transcriptTabStripPanel = New StackPanel() With {
                    .Orientation = Orientation.Horizontal
                }

                _transcriptTabStripScrollViewer = New ScrollViewer() With {
                    .HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden,
                    .VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
                    .Content = _transcriptTabStripPanel,
                    .CanContentScroll = False
                }

                _transcriptTabStripBorder = New Border() With {
                    .Visibility = Visibility.Collapsed,
                    .Child = _transcriptTabStripScrollViewer
                }

                Dim headerDockPanel = FindVisualAncestor(Of DockPanel)(WorkspacePaneHost.LblCurrentThread)
                Dim headerHostGrid = TryCast(If(headerDockPanel, Nothing)?.Parent, Grid)
                If headerDockPanel IsNot Nothing AndAlso headerHostGrid IsNot Nothing Then
                    Dim headerRow = Grid.GetRow(headerDockPanel)
                    If headerRow < 0 Then
                        headerRow = 0
                    End If

                    headerHostGrid.RowDefinitions.Insert(headerRow, New RowDefinition() With {.Height = GridLength.Auto})

                    For Each child As UIElement In headerHostGrid.Children
                        If child Is Nothing OrElse ReferenceEquals(child, _transcriptTabStripBorder) Then
                            Continue For
                        End If

                        Dim childRow = Grid.GetRow(child)
                        If childRow >= headerRow Then
                            Grid.SetRow(child, childRow + 1)
                        End If
                    Next

                    Dim headerMargin = headerDockPanel.Margin
                    _transcriptTabStripBorder.Margin = New Thickness(headerMargin.Left, headerMargin.Top, headerMargin.Right, 6)
                    headerDockPanel.Margin = New Thickness(headerMargin.Left, 0, headerMargin.Right, headerMargin.Bottom)

                    Grid.SetRow(_transcriptTabStripBorder, headerRow)
                    headerHostGrid.Children.Add(_transcriptTabStripBorder)
                Else
                    Dim titleStack = TryCast(WorkspacePaneHost.LblCurrentThread.Parent, StackPanel)
                    If titleStack IsNot Nothing Then
                        _transcriptTabStripBorder.Margin = New Thickness(0, 0, 0, 7)
                        titleStack.Children.Insert(0, _transcriptTabStripBorder)
                    End If
                End If
            End If
        End Sub

        Private Function EnsureTranscriptTabSurfaceHandle(threadId As String) As TranscriptTabSurfaceHandle
            EnsureTranscriptTabsUiInitialized()

            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return Nothing
            End If

            Dim existing As TranscriptTabSurfaceHandle = Nothing
            If _transcriptTabSurfacesByThreadId.TryGetValue(normalizedThreadId, existing) AndAlso existing IsNot Nothing Then
                Return existing
            End If

            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.LstTranscript Is Nothing Then
                Return Nothing
            End If

            Dim usePrimarySurface = _transcriptTabSurfacesByThreadId.Count = 0
            If Not usePrimarySurface AndAlso _transcriptSurfaceHostPanel Is Nothing Then
                Return Nothing
            End If

            Dim listBox = If(usePrimarySurface, WorkspacePaneHost.LstTranscript, CreateRetainedTranscriptSurfaceClone())
            If listBox Is Nothing Then
                Return Nothing
            End If

            If Not usePrimarySurface AndAlso _transcriptSurfaceHostPanel IsNot Nothing Then
                _transcriptSurfaceHostPanel.Children.Add(listBox)
            End If

            listBox.DataContext = _viewModel
            listBox.Visibility = Visibility.Collapsed

            Dim tabButton As Button = Nothing
            Dim tabCloseButton As Button = Nothing
            Dim tabChipBorder = CreateTranscriptTabChip(normalizedThreadId, tabButton, tabCloseButton)

            Dim handle As New TranscriptTabSurfaceHandle() With {
                .ThreadId = normalizedThreadId,
                .SurfaceListBox = listBox,
                .TabButton = tabButton,
                .TabChipBorder = tabChipBorder,
                .TabCloseButton = tabCloseButton,
                .IsPrimarySurface = usePrimarySurface
            }

            _transcriptTabSurfacesByThreadId(normalizedThreadId) = handle
            If _transcriptTabStripPanel IsNot Nothing AndAlso tabChipBorder IsNot Nothing Then
                _transcriptTabStripPanel.Children.Add(tabChipBorder)
            End If

            AttachTranscriptInteractionHandlers(listBox)
            RefreshTranscriptTabStripVisibility()
            UpdateTranscriptTabButtonCaption(handle)
            UpdateTranscriptTabButtonVisual(handle,
                                            StringComparer.Ordinal.Equals(normalizedThreadId, _activeTranscriptSurfaceThreadId))

            Return handle
        End Function

        Private Function CreateRetainedTranscriptSurfaceClone() As ListBox
            Dim source = If(WorkspacePaneHost, Nothing)?.LstTranscript
            If source Is Nothing Then
                Return Nothing
            End If

            Dim clone As New ListBox() With {
                .Background = source.Background,
                .BorderBrush = source.BorderBrush,
                .BorderThickness = source.BorderThickness,
                .FontFamily = source.FontFamily,
                .FontSize = source.FontSize,
                .HorizontalContentAlignment = source.HorizontalContentAlignment,
                .HorizontalAlignment = source.HorizontalAlignment,
                .VerticalAlignment = source.VerticalAlignment,
                .Margin = source.Margin,
                .Padding = source.Padding,
                .SelectionMode = source.SelectionMode,
                .ItemTemplate = source.ItemTemplate,
                .ItemContainerStyle = source.ItemContainerStyle,
                .ItemsPanel = source.ItemsPanel,
                .Visibility = Visibility.Collapsed
            }

            clone.SetValue(ScrollViewer.CanContentScrollProperty, source.GetValue(ScrollViewer.CanContentScrollProperty))
            clone.SetValue(VirtualizingPanel.IsVirtualizingProperty, source.GetValue(VirtualizingPanel.IsVirtualizingProperty))
            clone.SetValue(VirtualizingPanel.VirtualizationModeProperty, source.GetValue(VirtualizingPanel.VirtualizationModeProperty))
            clone.SetValue(VirtualizingPanel.ScrollUnitProperty, source.GetValue(VirtualizingPanel.ScrollUnitProperty))
            clone.SetValue(ScrollViewer.VerticalScrollBarVisibilityProperty, source.GetValue(ScrollViewer.VerticalScrollBarVisibilityProperty))
            clone.SetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty, source.GetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty))

            If _transcriptSurfaceHostPanel IsNot Nothing Then
                Grid.SetRow(clone, Grid.GetRow(source))
                Grid.SetColumn(clone, Grid.GetColumn(source))
                Grid.SetRowSpan(clone, Grid.GetRowSpan(source))
                Grid.SetColumnSpan(clone, Grid.GetColumnSpan(source))
                Panel.SetZIndex(clone, Panel.GetZIndex(source))
            End If

            For Each key As Object In source.Resources.Keys
                If clone.Resources.Contains(key) Then
                    Continue For
                End If

                clone.Resources.Add(key, source.Resources(key))
            Next

            Return clone
        End Function

        Private Function CreateTranscriptTabChip(threadId As String,
                                                 ByRef tabButton As Button,
                                                 ByRef closeButton As Button) As Border
            tabButton = New Button() With {
                .Tag = threadId,
                .Padding = New Thickness(10, 3, 6, 3),
                .MinHeight = 26,
                .FontSize = 12.5R,
                .Cursor = Cursors.Hand,
                .FocusVisualStyle = Nothing,
                .Background = Brushes.Transparent,
                .BorderBrush = Brushes.Transparent,
                .BorderThickness = New Thickness(0),
                .HorizontalContentAlignment = HorizontalAlignment.Left
            }

            Dim buttonBaseStyle = TryCast(TryFindResource("ButtonBaseStyle"), Style)
            If buttonBaseStyle IsNot Nothing Then
                tabButton.Style = buttonBaseStyle
            End If

            AddHandler tabButton.Click, AddressOf OnTranscriptTabButtonClick

            closeButton = New Button() With {
                .Tag = threadId,
                .Width = 18,
                .Height = 18,
                .MinWidth = 18,
                .MinHeight = 18,
                .Margin = New Thickness(0, 0, 6, 0),
                .Padding = New Thickness(0),
                .Cursor = Cursors.Hand,
                .FocusVisualStyle = Nothing,
                .Background = Brushes.Transparent,
                .BorderBrush = Brushes.Transparent,
                .BorderThickness = New Thickness(0),
                .Opacity = 0.0R,
                .IsHitTestVisible = False,
                .ToolTip = "Close tab",
                .Content = New TextBlock() With {
                    .Text = "x",
                    .FontSize = 11,
                    .HorizontalAlignment = HorizontalAlignment.Center,
                    .VerticalAlignment = VerticalAlignment.Center
                }
            }

            AddHandler closeButton.Click, AddressOf OnTranscriptTabCloseButtonClick

            Dim contentGrid As New Grid()
            contentGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(1, GridUnitType.Star)})
            contentGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = GridLength.Auto})
            Grid.SetColumn(tabButton, 0)
            Grid.SetColumn(closeButton, 1)
            contentGrid.Children.Add(tabButton)
            contentGrid.Children.Add(closeButton)

            Dim chipBorder As New Border() With {
                .Tag = threadId,
                .CornerRadius = New CornerRadius(8),
                .BorderThickness = New Thickness(1),
                .Margin = New Thickness(0, 0, 6, 0),
                .Child = contentGrid
            }

            AddHandler chipBorder.MouseEnter, AddressOf OnTranscriptTabChipMouseEnter
            AddHandler chipBorder.MouseLeave, AddressOf OnTranscriptTabChipMouseLeave

            Return chipBorder
        End Function

        Private Sub OnTranscriptTabButtonClick(sender As Object, e As RoutedEventArgs)
            Dim button = TryCast(sender, Button)
            If button Is Nothing Then
                Return
            End If

            Dim threadId = TryCast(button.Tag, String)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            Dim entry = FindVisibleThreadListEntryById(normalizedThreadId)
            If entry Is Nothing Then
                ShowStatus("This tab's thread is hidden by the current filter/search.", isError:=True, displayToast:=True)
                Return
            End If

            SelectThreadEntry(entry, suppressAutoLoad:=False)
        End Sub

        Private Sub OnTranscriptTabChipMouseEnter(sender As Object, e As MouseEventArgs)
            Dim chipBorder = TryCast(sender, Border)
            If chipBorder Is Nothing Then
                Return
            End If

            Dim handle = FindTranscriptTabSurfaceHandleByThreadId(TryCast(chipBorder.Tag, String))
            If handle Is Nothing Then
                Return
            End If

            UpdateTranscriptTabButtonVisual(handle,
                                            StringComparer.Ordinal.Equals(handle.ThreadId, _activeTranscriptSurfaceThreadId))
        End Sub

        Private Sub OnTranscriptTabChipMouseLeave(sender As Object, e As MouseEventArgs)
            Dim chipBorder = TryCast(sender, Border)
            If chipBorder Is Nothing Then
                Return
            End If

            Dim handle = FindTranscriptTabSurfaceHandleByThreadId(TryCast(chipBorder.Tag, String))
            If handle Is Nothing Then
                Return
            End If

            UpdateTranscriptTabButtonVisual(handle,
                                            StringComparer.Ordinal.Equals(handle.ThreadId, _activeTranscriptSurfaceThreadId))
        End Sub

        Private Async Sub OnTranscriptTabCloseButtonClick(sender As Object, e As RoutedEventArgs)
            e.Handled = True
            Dim closePerf = Stopwatch.StartNew()

            Dim button = TryCast(sender, Button)
            If button Is Nothing Then
                Return
            End If

            Dim normalizedThreadId = If(TryCast(button.Tag, String), String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            AppendProtocol("debug",
                           $"transcript_tab_perf event=close_begin thread={normalizedThreadId} active={StringComparer.Ordinal.Equals(normalizedThreadId, _activeTranscriptSurfaceThreadId)} openTabs={_transcriptTabSurfacesByThreadId.Count}")

            Dim isActiveTab = StringComparer.Ordinal.Equals(normalizedThreadId, _activeTranscriptSurfaceThreadId)
            Dim fallbackEntry As ThreadListEntry = Nothing
            If isActiveTab Then
                fallbackEntry = FindFirstVisibleTranscriptTabEntryExcluding(normalizedThreadId)
            End If

            RemoveRetainedTranscriptTabSurface(normalizedThreadId)
            _inactiveTranscriptDocumentsByThreadId.Remove(normalizedThreadId)

            If Not isActiveTab Then
                AppendProtocol("debug",
                               $"transcript_tab_perf event=close_complete thread={normalizedThreadId} mode=inactive_tab elapsedMs={closePerf.ElapsedMilliseconds} remainingTabs={_transcriptTabSurfacesByThreadId.Count}")
                Return
            End If

            If fallbackEntry IsNot Nothing Then
                SelectThreadEntry(fallbackEntry, suppressAutoLoad:=False)
                AppendProtocol("debug",
                               $"transcript_tab_perf event=close_complete thread={normalizedThreadId} mode=switch_to_fallback elapsedMs={closePerf.ElapsedMilliseconds} fallback={fallbackEntry.Id} remainingTabs={_transcriptTabSurfacesByThreadId.Count}")
                Return
            End If

            Await StartThreadAsync().ConfigureAwait(True)
            _inactiveTranscriptDocumentsByThreadId.Remove(normalizedThreadId)
            AppendProtocol("debug",
                           $"transcript_tab_perf event=close_complete thread={normalizedThreadId} mode=start_blank elapsedMs={closePerf.ElapsedMilliseconds} remainingTabs={_transcriptTabSurfacesByThreadId.Count}")
        End Sub

        Private Function FindFirstVisibleTranscriptTabEntryExcluding(threadId As String) As ThreadListEntry
            Dim excludedThreadId = If(threadId, String.Empty).Trim()

            For Each kvp In _transcriptTabSurfacesByThreadId
                Dim candidateHandle = kvp.Value
                If candidateHandle Is Nothing OrElse String.IsNullOrWhiteSpace(candidateHandle.ThreadId) Then
                    Continue For
                End If

                If StringComparer.Ordinal.Equals(candidateHandle.ThreadId, excludedThreadId) Then
                    Continue For
                End If

                Dim entry = FindVisibleThreadListEntryById(candidateHandle.ThreadId)
                If entry IsNot Nothing Then
                    Return entry
                End If
            Next

            Return Nothing
        End Function

        Private Function FindTranscriptTabSurfaceHandleByThreadId(threadId As String) As TranscriptTabSurfaceHandle
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return Nothing
            End If

            Dim handle As TranscriptTabSurfaceHandle = Nothing
            If _transcriptTabSurfacesByThreadId.TryGetValue(normalizedThreadId, handle) Then
                Return handle
            End If

            Return Nothing
        End Function

        Private Sub EnsureTranscriptTabSurfaceActivatedForThread(threadId As String)
            EnsureTranscriptTabsUiInitialized()

            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.LstTranscript Is Nothing Then
                Return
            End If

            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                ActivateBlankTranscriptSurface()
                Return
            End If

            Dim handle = EnsureTranscriptTabSurfaceHandle(normalizedThreadId)
            If handle Is Nothing OrElse handle.SurfaceListBox Is Nothing Then
                Return
            End If

            handle.SurfaceListBox.DataContext = _viewModel
            handle.SurfaceListBox.ItemsSource = _viewModel.TranscriptPanel.Items

            For Each kvp In _transcriptTabSurfacesByThreadId
                Dim candidate = kvp.Value
                If candidate Is Nothing OrElse candidate.SurfaceListBox Is Nothing Then
                    Continue For
                End If

                Dim isVisible = StringComparer.Ordinal.Equals(candidate.ThreadId, normalizedThreadId)
                candidate.SurfaceListBox.Visibility = If(isVisible, Visibility.Visible, Visibility.Collapsed)
                candidate.SurfaceListBox.IsHitTestVisible = isVisible
            Next

            If Not HandleOwnsPrimaryTranscriptSurface(normalizedThreadId) Then
                WorkspacePaneHost.LstTranscript.Visibility = Visibility.Collapsed
                WorkspacePaneHost.LstTranscript.IsHitTestVisible = False
            End If

            _activeTranscriptSurfaceListBox = handle.SurfaceListBox
            _activeTranscriptSurfaceThreadId = normalizedThreadId
            _transcriptScrollViewer = Nothing

            UpdateTranscriptTabButtonCaptions()
            UpdateTranscriptTabButtonVisuals()
            RefreshTranscriptTabStripVisibility()
            UpdateWorkspaceEmptyStateVisibility()
        End Sub

        Private Sub ActivateBlankTranscriptSurface()
            EnsureTranscriptTabsUiInitialized()

            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.LstTranscript Is Nothing Then
                Return
            End If

            WorkspacePaneHost.LstTranscript.DataContext = _viewModel
            WorkspacePaneHost.LstTranscript.ItemsSource = _viewModel.TranscriptPanel.Items
            WorkspacePaneHost.LstTranscript.Visibility = Visibility.Visible
            WorkspacePaneHost.LstTranscript.IsHitTestVisible = True

            For Each kvp In _transcriptTabSurfacesByThreadId
                Dim candidate = kvp.Value
                If candidate Is Nothing OrElse candidate.SurfaceListBox Is Nothing Then
                    Continue For
                End If

                candidate.SurfaceListBox.Visibility = Visibility.Collapsed
                candidate.SurfaceListBox.IsHitTestVisible = False
            Next

            _activeTranscriptSurfaceListBox = WorkspacePaneHost.LstTranscript
            _activeTranscriptSurfaceThreadId = String.Empty
            _transcriptScrollViewer = Nothing

            UpdateTranscriptTabButtonVisuals()
            RefreshTranscriptTabStripVisibility()
            UpdateWorkspaceEmptyStateVisibility()
        End Sub

        Private Function HandleOwnsPrimaryTranscriptSurface(threadId As String) As Boolean
            Dim handle As TranscriptTabSurfaceHandle = Nothing
            If Not _transcriptTabSurfacesByThreadId.TryGetValue(If(threadId, String.Empty), handle) OrElse handle Is Nothing Then
                Return False
            End If

            Return handle.IsPrimarySurface AndAlso ReferenceEquals(handle.SurfaceListBox, If(WorkspacePaneHost, Nothing)?.LstTranscript)
        End Function

        Private Function HasRetainedTranscriptTabSurface(threadId As String) As Boolean
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return False
            End If

            Dim handle As TranscriptTabSurfaceHandle = Nothing
            Return _transcriptTabSurfacesByThreadId.TryGetValue(normalizedThreadId, handle) AndAlso
                   handle IsNot Nothing AndAlso
                   handle.SurfaceListBox IsNot Nothing
        End Function

        Private Sub QueueTranscriptTabSurfaceRetireForIdle(handle As TranscriptTabSurfaceHandle)
            If handle Is Nothing OrElse handle.SurfaceListBox Is Nothing Then
                Return
            End If

            Dim listBox = handle.SurfaceListBox
            listBox.Visibility = Visibility.Collapsed
            listBox.IsHitTestVisible = False

            _transcriptTabSurfaceRetireQueue.Enqueue(New TranscriptTabSurfaceRetireWorkItem() With {
                .ThreadId = If(handle.ThreadId, String.Empty).Trim(),
                .SurfaceListBox = listBox,
                .IsPrimarySurface = handle.IsPrimarySurface
            })

            AppendProtocol("debug",
                           $"transcript_tab_perf event=surface_retire_queued thread={If(handle.ThreadId, String.Empty)} queueCount={_transcriptTabSurfaceRetireQueue.Count}")

            ScheduleTranscriptTabSurfaceRetireDrain()
        End Sub

        Private Sub ScheduleTranscriptTabSurfaceRetireDrain()
            If _transcriptTabSurfaceRetireDrainScheduled Then
                Return
            End If

            _transcriptTabSurfaceRetireDrainScheduled = True
            Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle,
                                   New Action(AddressOf DrainTranscriptTabSurfaceRetireQueue))
        End Sub

        Private Sub DrainTranscriptTabSurfaceRetireQueue()
            _transcriptTabSurfaceRetireDrainScheduled = False

            If _transcriptTabSurfaceRetireQueue.Count = 0 Then
                Return
            End If

            Dim retirePerf = Stopwatch.StartNew()
            Dim processed = 0
            Const maxItemsPerIdlePass As Integer = 1
            Const maxIdleBudgetMs As Long = 12

            Do While _transcriptTabSurfaceRetireQueue.Count > 0 AndAlso
                     processed < maxItemsPerIdlePass AndAlso
                     retirePerf.ElapsedMilliseconds <= maxIdleBudgetMs
                Dim workItem = _transcriptTabSurfaceRetireQueue.Dequeue()
                If workItem Is Nothing OrElse workItem.SurfaceListBox Is Nothing Then
                    Continue Do
                End If

                DisposeRetiredTranscriptSurfaceWorkItem(workItem)

                processed += 1

                AppendProtocol("debug",
                               $"transcript_tab_perf event=surface_retire_disposed thread={If(workItem.ThreadId, String.Empty)} queuedMs={CLng((DateTimeOffset.UtcNow - workItem.QueuedUtc).TotalMilliseconds)} remaining={_transcriptTabSurfaceRetireQueue.Count}")
            Loop

            If _transcriptTabSurfaceRetireQueue.Count > 0 Then
                ScheduleTranscriptTabSurfaceRetireDrain()
            End If
        End Sub

        Private Sub DisposeRetiredTranscriptSurfaceWorkItem(workItem As TranscriptTabSurfaceRetireWorkItem)
            If workItem Is Nothing OrElse workItem.SurfaceListBox Is Nothing Then
                Return
            End If

            Dim listBox = workItem.SurfaceListBox

            If Not workItem.IsPrimarySurface AndAlso _transcriptSurfaceHostPanel IsNot Nothing Then
                _transcriptSurfaceHostPanel.Children.Remove(listBox)
            End If

            listBox.ItemsSource = Nothing
            listBox.DataContext = Nothing

            For i = _transcriptInteractionHandlersAttached.Count - 1 To 0 Step -1
                If ReferenceEquals(_transcriptInteractionHandlersAttached(i), listBox) Then
                    _transcriptInteractionHandlersAttached.RemoveAt(i)
                End If
            Next

            For i = _transcriptChunkScrollChangedHandlerAttachedLists.Count - 1 To 0 Step -1
                If ReferenceEquals(_transcriptChunkScrollChangedHandlerAttachedLists(i), listBox) Then
                    _transcriptChunkScrollChangedHandlerAttachedLists.RemoveAt(i)
                End If
            Next
        End Sub

        Private Sub FlushQueuedTranscriptTabSurfaceRetiresImmediately()
            If _transcriptTabSurfaceRetireQueue.Count = 0 Then
                _transcriptTabSurfaceRetireDrainScheduled = False
                Return
            End If

            Do While _transcriptTabSurfaceRetireQueue.Count > 0
                DisposeRetiredTranscriptSurfaceWorkItem(_transcriptTabSurfaceRetireQueue.Dequeue())
            Loop

            _transcriptTabSurfaceRetireDrainScheduled = False
        End Sub

        Private Sub RemoveRetainedTranscriptTabSurface(threadId As String)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            Dim handle As TranscriptTabSurfaceHandle = Nothing
            If Not _transcriptTabSurfacesByThreadId.TryGetValue(normalizedThreadId, handle) OrElse handle Is Nothing Then
                Return
            End If

            _transcriptTabSurfacesByThreadId.Remove(normalizedThreadId)

            If handle.TabChipBorder IsNot Nothing AndAlso _transcriptTabStripPanel IsNot Nothing Then
                _transcriptTabStripPanel.Children.Remove(handle.TabChipBorder)
            End If

            If handle.SurfaceListBox IsNot Nothing Then
                If handle.IsPrimarySurface AndAlso WorkspacePaneHost IsNot Nothing AndAlso ReferenceEquals(handle.SurfaceListBox, WorkspacePaneHost.LstTranscript) Then
                    handle.SurfaceListBox.Visibility = Visibility.Collapsed
                    handle.SurfaceListBox.IsHitTestVisible = False
                Else
                    QueueTranscriptTabSurfaceRetireForIdle(handle)
                End If
            End If

            If StringComparer.Ordinal.Equals(_activeTranscriptSurfaceThreadId, normalizedThreadId) Then
                _activeTranscriptSurfaceThreadId = String.Empty
                _activeTranscriptSurfaceListBox = If(WorkspacePaneHost, Nothing)?.LstTranscript
            End If

            RefreshTranscriptTabStripVisibility()
            UpdateTranscriptTabButtonVisuals()
        End Sub

        Private Sub ClearRetainedTranscriptTabSurfaces()
            EnsureTranscriptTabsUiInitialized()

            Dim surfaceHandles As New List(Of TranscriptTabSurfaceHandle)(_transcriptTabSurfacesByThreadId.Values)
            _transcriptTabSurfacesByThreadId.Clear()
            FlushQueuedTranscriptTabSurfaceRetiresImmediately()

            For Each handle In surfaceHandles
                If handle Is Nothing Then
                    Continue For
                End If

                If handle.TabChipBorder IsNot Nothing AndAlso _transcriptTabStripPanel IsNot Nothing Then
                    _transcriptTabStripPanel.Children.Remove(handle.TabChipBorder)
                End If

                If handle.SurfaceListBox Is Nothing Then
                    Continue For
                End If

                If WorkspacePaneHost IsNot Nothing AndAlso ReferenceEquals(handle.SurfaceListBox, WorkspacePaneHost.LstTranscript) Then
                    handle.SurfaceListBox.Visibility = Visibility.Visible
                    handle.SurfaceListBox.IsHitTestVisible = True
                    handle.SurfaceListBox.DataContext = _viewModel
                    handle.SurfaceListBox.ItemsSource = _viewModel.TranscriptPanel.Items
                ElseIf _transcriptSurfaceHostPanel IsNot Nothing Then
                    _transcriptSurfaceHostPanel.Children.Remove(handle.SurfaceListBox)
                    handle.SurfaceListBox.ItemsSource = Nothing
                    handle.SurfaceListBox.DataContext = Nothing
                End If
            Next

            Dim primaryListBox = If(WorkspacePaneHost, Nothing)?.LstTranscript
            For i = _transcriptInteractionHandlersAttached.Count - 1 To 0 Step -1
                If primaryListBox Is Nothing OrElse
                   Not ReferenceEquals(_transcriptInteractionHandlersAttached(i), primaryListBox) Then
                    _transcriptInteractionHandlersAttached.RemoveAt(i)
                End If
            Next

            For i = _transcriptChunkScrollChangedHandlerAttachedLists.Count - 1 To 0 Step -1
                If primaryListBox Is Nothing OrElse
                   Not ReferenceEquals(_transcriptChunkScrollChangedHandlerAttachedLists(i), primaryListBox) Then
                    _transcriptChunkScrollChangedHandlerAttachedLists.RemoveAt(i)
                End If
            Next

            If primaryListBox IsNot Nothing Then
                AttachTranscriptInteractionHandlers(primaryListBox)
                _activeTranscriptSurfaceListBox = WorkspacePaneHost.LstTranscript
                primaryListBox.IsHitTestVisible = True
            Else
                _activeTranscriptSurfaceListBox = Nothing
            End If

            _activeTranscriptSurfaceThreadId = String.Empty
            _transcriptScrollViewer = Nothing
            RefreshTranscriptTabStripVisibility()
        End Sub

        Private Sub RefreshTranscriptTabStripVisibility()
            If _transcriptTabStripBorder Is Nothing Then
                Return
            End If

            _transcriptTabStripBorder.Visibility = If(_transcriptTabSurfacesByThreadId.Count > 0,
                                                      Visibility.Visible,
                                                      Visibility.Collapsed)
        End Sub

        Private Sub UpdateTranscriptTabButtonCaptions()
            For Each kvp In _transcriptTabSurfacesByThreadId
                UpdateTranscriptTabButtonCaption(kvp.Value)
            Next
        End Sub

        Private Sub UpdateTranscriptTabButtonCaption(handle As TranscriptTabSurfaceHandle)
            If handle Is Nothing OrElse handle.TabButton Is Nothing Then
                Return
            End If

            Dim threadId = If(handle.ThreadId, String.Empty).Trim()
            Dim text = ResolveTranscriptTabCaption(threadId)
            handle.TabButton.Content = text
            handle.TabButton.ToolTip = threadId
        End Sub

        Private Function ResolveTranscriptTabCaption(threadId As String) As String
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return "New thread"
            End If

            If StringComparer.Ordinal.Equals(normalizedThreadId, GetVisibleThreadId()) Then
                Dim currentLabel = If(_viewModel, Nothing)?.CurrentThreadText
                If Not String.IsNullOrWhiteSpace(currentLabel) AndAlso
                   Not StringComparer.Ordinal.Equals(currentLabel, "New thread") Then
                    Return CompactTranscriptTabCaption(currentLabel)
                End If
            End If

            Dim entry = FindVisibleThreadListEntryById(normalizedThreadId)
            If entry IsNot Nothing Then
                Dim preview = If(entry.Preview, String.Empty).Trim()
                If Not String.IsNullOrWhiteSpace(preview) Then
                    Return CompactTranscriptTabCaption(preview)
                End If
            End If

            Return CompactTranscriptTabCaption(normalizedThreadId)
        End Function

        Private Shared Function CompactTranscriptTabCaption(value As String) As String
            Dim text = If(value, String.Empty).Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            If String.IsNullOrWhiteSpace(text) Then
                Return "Thread"
            End If

            Const maxLen As Integer = 26
            If text.Length <= maxLen Then
                Return text
            End If

            Return text.Substring(0, maxLen - 3).TrimEnd() & "..."
        End Function

        Private Sub UpdateTranscriptTabButtonVisuals()
            For Each kvp In _transcriptTabSurfacesByThreadId
                UpdateTranscriptTabButtonVisual(kvp.Value,
                                               StringComparer.Ordinal.Equals(kvp.Key, _activeTranscriptSurfaceThreadId))
            Next
        End Sub

        Private Sub UpdateTranscriptTabButtonVisual(handle As TranscriptTabSurfaceHandle, isActive As Boolean)
            If handle Is Nothing OrElse handle.TabButton Is Nothing Then
                Return
            End If

            Dim button = handle.TabButton
            Dim chipBorder = handle.TabChipBorder
            Dim background = TryCast(TryFindResource(If(isActive, "SurfaceMutedBrush", "SurfaceBrush")), Brush)
            Dim border = TryCast(TryFindResource(If(isActive, "AccentGlowBrush", "BorderBrush")), Brush)
            Dim foreground = TryCast(TryFindResource(If(isActive, "TextPrimaryBrush", "TextSecondaryBrush")), Brush)

            If chipBorder IsNot Nothing Then
                chipBorder.Background = If(background, Brushes.Transparent)
                chipBorder.BorderBrush = If(border, Brushes.Transparent)
                chipBorder.BorderThickness = New Thickness(1)
                chipBorder.Opacity = If(isActive, 1.0R, 0.94R)
            End If

            button.Background = Brushes.Transparent
            button.BorderBrush = Brushes.Transparent
            button.BorderThickness = New Thickness(0)
            button.Foreground = If(foreground, Brushes.Black)
            button.FontWeight = If(isActive, FontWeights.SemiBold, FontWeights.Normal)
            button.Opacity = 1.0R

            UpdateTranscriptTabCloseButtonVisual(handle, isActive, foreground)
        End Sub

        Private Sub UpdateTranscriptTabCloseButtonVisual(handle As TranscriptTabSurfaceHandle,
                                                         isActive As Boolean,
                                                         foreground As Brush)
            If handle Is Nothing OrElse handle.TabCloseButton Is Nothing Then
                Return
            End If

            Dim closeButton = handle.TabCloseButton
            Dim chipBorder = handle.TabChipBorder
            Dim showClose = chipBorder IsNot Nothing AndAlso chipBorder.IsMouseOver

            closeButton.Opacity = If(showClose, 1.0R, 0.0R)
            closeButton.IsHitTestVisible = showClose
            closeButton.Foreground = If(foreground, Brushes.Black)

            Dim closeGlyph = TryCast(closeButton.Content, TextBlock)
            If closeGlyph IsNot Nothing Then
                closeGlyph.Foreground = If(foreground, Brushes.Black)
                closeGlyph.FontWeight = If(isActive, FontWeights.SemiBold, FontWeights.Normal)
                closeGlyph.Opacity = If(showClose, 0.9R, 0.0R)
            End If
        End Sub

        Private Sub RefreshActiveTranscriptTabCaption()
            If String.IsNullOrWhiteSpace(_activeTranscriptSurfaceThreadId) Then
                Return
            End If

            Dim handle As TranscriptTabSurfaceHandle = Nothing
            If _transcriptTabSurfacesByThreadId.TryGetValue(_activeTranscriptSurfaceThreadId, handle) Then
                UpdateTranscriptTabButtonCaption(handle)
                UpdateTranscriptTabButtonVisual(handle, True)
            End If
        End Sub
    End Class
End Namespace
