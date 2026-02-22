Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports CodexNativeAgent.Services
Imports CodexNativeAgent.Ui.Coordinators
Imports CodexNativeAgent.Ui.ViewModels.Transcript
Imports CodexNativeAgent.Ui.ViewModels.Threads

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private NotInheritable Class ThreadSelectionLoadRequest
            Public Property ThreadId As String = String.Empty
            Public Property LoadVersion As Integer
            Public Property CancellationSource As CancellationTokenSource

            Public ReadOnly Property CancellationToken As CancellationToken
                Get
                    If CancellationSource Is Nothing Then
                        Return CancellationToken.None
                    End If

                    Return CancellationSource.Token
                End Get
            End Property
        End Class

        Private NotInheritable Class ThreadSelectionLoadPayload
            Public Property ThreadObject As JsonObject
            Public Property HasTurns As Boolean
            Public Property TranscriptSnapshot As ThreadTranscriptSnapshotData
        End Class

        Private NotInheritable Class ThreadTranscriptSnapshotData
            Public Property RawText As String = String.Empty
            Public ReadOnly Property DisplayEntries As New List(Of TranscriptEntryDescriptor)()
        End Class

        Private Function StartThreadAsync() As Task
            CancelActiveThreadSelectionLoad()
            ResetThreadSelectionLoadUiState(hideTranscriptLoader:=True)
            _viewModel.TranscriptPanel.ClearTranscript()
            _currentThreadId = String.Empty
            _currentTurnId = String.Empty
            SetPendingNewThreadFirstPromptSelectionActive(True, clearThreadSelection:=True)
            UpdateThreadTurnLabels()
            RefreshControlStates()
            ShowStatus("New thread ready. Send your first instruction.")
            Return Task.CompletedTask
        End Function

        Private Function CanBeginThreadsRefresh() As Boolean
            Return IsClientRunning()
        End Function

        Private Sub BeginThreadsRefreshUi(Optional silent As Boolean = False)
            _threadsLoading = True
            UpdateThreadsPanelInteractionState()
            _viewModel.ThreadsPanel.BeginThreadsRefreshState()
            UpdateThreadsStateLabel(VisibleThreadCount())
            RefreshControlStates()

            If Not silent Then
                ShowStatus("Loading threads...")
            End If
        End Sub

        Private Sub ApplyThreadSummariesToEntries(summaries As IReadOnlyList(Of ThreadSummary))
            _threadEntries.Clear()

            If summaries Is Nothing Then
                Return
            End If

            For Each summary In summaries
                If summary Is Nothing Then
                    Continue For
                End If

                Dim lastActiveText = If(String.IsNullOrWhiteSpace(summary.LastActiveText),
                                        summary.UpdatedAtText,
                                        summary.LastActiveText)
                Dim lastActiveSortTimestamp = summary.LastActiveSortValue
                If lastActiveSortTimestamp = Long.MinValue Then
                    lastActiveSortTimestamp = summary.UpdatedSortValue
                End If

                _threadEntries.Add(New ThreadListEntry() With {
                    .Id = summary.Id,
                    .Preview = summary.Preview,
                    .LastActiveAt = lastActiveText,
                    .LastActiveSortTimestamp = lastActiveSortTimestamp,
                    .Cwd = summary.Cwd,
                    .IsArchived = summary.IsArchived
                })
            Next
        End Sub

        Private Sub CompleteThreadsRefreshUi(Optional silent As Boolean = False)
            _threadsLoadedAtLeastOnce = True
            _viewModel.ThreadsPanel.CompleteThreadsRefreshState(_threadEntries.Count)
            ApplyThreadFiltersAndSort()

            If Not silent Then
                AppendSystemMessage($"Loaded {_threadEntries.Count} thread(s).")
                ShowStatus($"Loaded {_threadEntries.Count} thread(s).")
            End If
        End Sub

        Private Sub FailThreadsRefreshUi(ex As Exception)
            Dim message = If(ex Is Nothing, "Unknown error", ex.Message)
            _viewModel.ThreadsPanel.FailThreadsRefreshState(message)
            UpdateThreadsStateLabel(VisibleThreadCount())
            ShowStatus($"Could not load threads: {message}", isError:=True, displayToast:=True)
        End Sub

        Private Sub FinalizeThreadsRefreshUi()
            _threadsLoading = False
            UpdateThreadsPanelInteractionState()
            UpdateThreadsStateLabel(VisibleThreadCount())
            RefreshControlStates()
        End Sub

        Private Async Function RefreshThreadsCoreAsync(silent As Boolean) As Task
            Await _threadWorkflowCoordinator.RunRefreshThreadsAsync(
                AddressOf CanBeginThreadsRefresh,
                Sub() BeginThreadsRefreshUi(silent),
                _viewModel.ThreadsPanel.ShowArchived,
                _viewModel.ThreadsPanel.FilterByWorkingDir,
                AddressOf EffectiveThreadWorkingDirectory,
                Function(showArchived, cwd, token)
                    Return _threadService.ListThreadsAsync(showArchived, cwd, token)
                End Function,
                AddressOf ApplyThreadSummariesToEntries,
                Sub() CompleteThreadsRefreshUi(silent),
                AddressOf FailThreadsRefreshUi,
                AddressOf FinalizeThreadsRefreshUi)
        End Function

        Private Async Function RefreshThreadsAsync() As Task
            Await RefreshThreadsCoreAsync(False)
        End Function

        Private Function TryPrepareAutoLoadThreadSelection(selected As ThreadListEntry,
                                                           forceReload As Boolean,
                                                           ByRef selectedThreadId As String) As Boolean
            selectedThreadId = String.Empty
            If selected Is Nothing OrElse String.IsNullOrWhiteSpace(selected.Id) Then
                Return False
            End If

            UpdateThreadsPanelInteractionState()
            If Not _viewModel.ThreadsPanel.CanAutoLoadSelection Then
                Return False
            End If

            selectedThreadId = selected.Id.Trim()
            If String.IsNullOrWhiteSpace(selectedThreadId) Then
                Return False
            End If

            If Not forceReload AndAlso
               StringComparer.Ordinal.Equals(selectedThreadId, _currentThreadId) AndAlso
               Not _threadContentLoading Then
                Return False
            End If

            Return True
        End Function

        Private Function BeginThreadSelectionLoadRequest(selectedThreadId As String) As ThreadSelectionLoadRequest
            Dim request As New ThreadSelectionLoadRequest() With {
                .ThreadId = If(selectedThreadId, String.Empty).Trim()
            }

            request.LoadVersion = BeginThreadSelectionLoadUiState(request.ThreadId)
            CancelActiveThreadSelectionLoad()

            Dim threadLoadCts As New CancellationTokenSource()
            request.CancellationSource = threadLoadCts
            _threadSelectionLoadCts = threadLoadCts

            SetTranscriptLoadingState(True, "Loading selected thread...")
            RefreshControlStates()
            ShowStatus("Loading selected thread...")
            Return request
        End Function

        Private Async Function LoadThreadSelectionPayloadAsync(threadId As String,
                                                               cancellationToken As CancellationToken) As Task(Of ThreadSelectionLoadPayload)
            Dim threadObject = Await _threadService.ResumeThreadAsync(threadId,
                                                                      New ThreadRequestOptions(),
                                                                      cancellationToken).ConfigureAwait(False)
            If Not ThreadObjectHasTurns(threadObject) Then
                threadObject = Await _threadService.ReadThreadAsync(threadId,
                                                                    includeTurns:=True,
                                                                    cancellationToken:=cancellationToken).ConfigureAwait(False)
            End If
            cancellationToken.ThrowIfCancellationRequested()

            Dim hasTurns = ThreadObjectHasTurns(threadObject)
            Dim transcriptSnapshot = Await Task.Run(Function() BuildThreadTranscriptSnapshot(threadObject), cancellationToken).ConfigureAwait(False)
            cancellationToken.ThrowIfCancellationRequested()

            Return New ThreadSelectionLoadPayload() With {
                .ThreadObject = threadObject,
                .HasTurns = hasTurns,
                .TranscriptSnapshot = transcriptSnapshot
            }
        End Function

        Private Async Function ApplyThreadSelectionPayloadUiAsync(request As ThreadSelectionLoadRequest,
                                                                  payload As ThreadSelectionLoadPayload) As Task
            If request Is Nothing OrElse payload Is Nothing Then
                Return
            End If

            Dim cancellationToken = request.CancellationToken
            Await RunOnUiAsync(
                Function()
                    If cancellationToken.IsCancellationRequested OrElse
                       Not IsCurrentThreadSelectionLoadUiState(request.LoadVersion, request.ThreadId) Then
                        Return Task.CompletedTask
                    End If

                    Dim selectedNow = TryCast(_viewModel.ThreadsPanel.SelectedListItem, ThreadListEntry)
                    If selectedNow Is Nothing OrElse Not StringComparer.Ordinal.Equals(selectedNow.Id, request.ThreadId) Then
                        Return Task.CompletedTask
                    End If

                    ApplyCurrentThreadFromThreadObject(payload.ThreadObject)
                    ApplyThreadTranscriptSnapshot(payload.TranscriptSnapshot, payload.HasTurns)
                    AppendSystemMessage($"Loaded thread {_currentThreadId} from history.")
                    ShowStatus($"Loaded thread {_currentThreadId}.")
                    Return Task.CompletedTask
                End Function).ConfigureAwait(False)
        End Function

        Private Sub HandleThreadSelectionLoadFailureUi(request As ThreadSelectionLoadRequest, ex As Exception)
            If request Is Nothing OrElse ex Is Nothing Then
                Return
            End If

            _viewModel.ThreadsPanel.RecordThreadSelectionLoadError(request.LoadVersion, request.ThreadId, ex.Message)
            If IsCurrentThreadSelectionLoadUiState(request.LoadVersion, request.ThreadId) Then
                ShowStatus($"Could not load thread {request.ThreadId}: {ex.Message}", isError:=True, displayToast:=True)
                AppendTranscript("system", $"Could not load thread {request.ThreadId}: {ex.Message}")
            End If
        End Sub

        Private Sub FinalizeThreadSelectionLoadRequestUi(request As ThreadSelectionLoadRequest)
            If request Is Nothing Then
                Return
            End If

            If TryCompleteThreadSelectionLoadUiState(request.LoadVersion) Then
                SetTranscriptLoadingState(False)
                RefreshControlStates()
            End If
        End Sub

        Private Sub DisposeThreadSelectionLoadRequest(request As ThreadSelectionLoadRequest)
            If request Is Nothing Then
                Return
            End If

            If _threadSelectionLoadCts Is request.CancellationSource Then
                _threadSelectionLoadCts = Nothing
            End If

            If request.CancellationSource IsNot Nothing Then
                request.CancellationSource.Dispose()
                request.CancellationSource = Nothing
            End If
        End Sub

        Private Async Function AutoLoadThreadSelectionAsync(selected As ThreadListEntry,
                                                            Optional forceReload As Boolean = False) As Task
            Await _threadWorkflowCoordinator.RunAutoLoadThreadSelectionAsync(
                selected,
                forceReload,
                Function(entry, force)
                    Dim selectedThreadId As String = String.Empty
                    If Not TryPrepareAutoLoadThreadSelection(entry, force, selectedThreadId) Then
                        Return String.Empty
                    End If

                    Return selectedThreadId
                End Function,
                Function(selectedThreadId)
                    Return CType(BeginThreadSelectionLoadRequest(selectedThreadId), Object)
                End Function,
                Function(requestObject)
                    Dim request = DirectCast(requestObject, ThreadSelectionLoadRequest)
                    Return request.ThreadId
                End Function,
                Function(requestObject)
                    Dim request = DirectCast(requestObject, ThreadSelectionLoadRequest)
                    Return request.CancellationToken
                End Function,
                Async Function(threadId, cancellationToken)
                    Dim payload = Await LoadThreadSelectionPayloadAsync(threadId, cancellationToken).ConfigureAwait(False)
                    Return CType(payload, Object)
                End Function,
                Async Function(requestObject, payloadObject)
                    Dim request = DirectCast(requestObject, ThreadSelectionLoadRequest)
                    Dim payload = DirectCast(payloadObject, ThreadSelectionLoadPayload)
                    Await ApplyThreadSelectionPayloadUiAsync(request, payload).ConfigureAwait(False)
                End Function,
                Sub(requestObject, ex)
                    Dim request = DirectCast(requestObject, ThreadSelectionLoadRequest)
                    RunOnUi(
                        Sub() HandleThreadSelectionLoadFailureUi(request, ex))
                End Sub,
                Sub(requestObject)
                    Dim request = DirectCast(requestObject, ThreadSelectionLoadRequest)
                    RunOnUi(
                        Sub() FinalizeThreadSelectionLoadRequestUi(request))
                End Sub,
                Sub(requestObject)
                    Dim request = DirectCast(requestObject, ThreadSelectionLoadRequest)
                    DisposeThreadSelectionLoadRequest(request)
                End Sub)
        End Function

        Private Sub CancelActiveThreadSelectionLoad()
            Dim cts = _threadSelectionLoadCts
            _threadSelectionLoadCts = Nothing
            If cts Is Nothing Then
                Return
            End If

            Try
                cts.Cancel()
            Catch
            End Try

            cts.Dispose()
        End Sub

        Private Sub SetTranscriptLoadingState(isLoading As Boolean, Optional loadingText As String = "Loading thread...")
            _viewModel.TranscriptPanel.LoadingOverlayVisibility = If(isLoading, Visibility.Visible, Visibility.Collapsed)
            _viewModel.TranscriptPanel.LoadingText = If(isLoading, loadingText, "Loading thread...")
            UpdateWorkspaceHintOverlayVisibility()
            UpdateWorkspaceEmptyStateVisibility()
        End Sub

        Private Sub OnThreadsPreviewMouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)
            _threadContextTarget = Nothing
            _threadGroupContextTarget = Nothing

            Dim source = TryCast(e.OriginalSource, DependencyObject)
            Dim container = FindVisualAncestor(Of ListBoxItem)(source)
            If container Is Nothing Then
                e.Handled = True
                Return
            End If

            Dim header = TryCast(container.DataContext, ThreadGroupHeaderEntry)
            If header IsNot Nothing Then
                _threadGroupContextTarget = header
                e.Handled = True
                PrepareThreadGroupContextMenu(header)
                SidebarPaneHost.ThreadItemContextMenu.PlacementTarget = container
                SidebarPaneHost.ThreadItemContextMenu.IsOpen = True
                Return
            End If

            Dim entry = TryCast(container.DataContext, ThreadListEntry)
            If entry IsNot Nothing Then
                _threadContextTarget = entry
                e.Handled = True
                PrepareThreadContextMenu(entry)
                SidebarPaneHost.ThreadItemContextMenu.PlacementTarget = container
                SidebarPaneHost.ThreadItemContextMenu.IsOpen = True
                Return
            End If

            e.Handled = True
        End Sub

        Private Sub OnThreadsPreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
            If e Is Nothing OrElse e.ChangedButton <> MouseButton.Left Then
                Return
            End If

            Dim source = TryCast(e.OriginalSource, DependencyObject)
            If source Is Nothing Then
                Return
            End If

            Dim container = FindVisualAncestor(Of ListBoxItem)(source)
            If container Is Nothing Then
                Return
            End If

            Dim header = TryCast(container.DataContext, ThreadGroupHeaderEntry)
            If header Is Nothing OrElse String.IsNullOrWhiteSpace(header.GroupKey) Then
                Return
            End If

            e.Handled = True
            ToggleThreadProjectGroupExpansion(header.GroupKey)
            ApplyThreadFiltersAndSort()
        End Sub

        Private Sub OnThreadsContextMenuOpening(sender As Object, e As ContextMenuEventArgs)
            Dim headerTarget = ResolveContextThreadGroupEntry()
            If headerTarget IsNot Nothing Then
                PrepareThreadGroupContextMenu(headerTarget)
                Return
            End If

            Dim target = ResolveContextThreadEntry()
            If target Is Nothing Then
                e.Handled = True
                Return
            End If

            PrepareThreadContextMenu(target)
        End Sub

        Private Sub PrepareThreadContextMenu(target As ThreadListEntry)
            If target Is Nothing Then
                Return
            End If

            UpdateThreadsPanelInteractionState()
            _viewModel.ThreadsPanel.ConfigureThreadContextMenuForThread(_viewModel.ThreadsPanel.CanRunThreadContextActions,
                                                                        target.IsArchived)
        End Sub

        Private Sub PrepareThreadGroupContextMenu(target As ThreadGroupHeaderEntry)
            If target Is Nothing Then
                Return
            End If

            UpdateThreadsPanelInteractionState()
            Dim canStartHere = _viewModel.ThreadsPanel.CanRunThreadContextActions AndAlso
                               Not String.IsNullOrWhiteSpace(target.ProjectPath)
            _viewModel.ThreadsPanel.ConfigureThreadContextMenuForGroup(canStartHere)
        End Sub

        Private Sub OnSelectThreadFromContextMenuClick(sender As Object, e As RoutedEventArgs)
            SelectThreadFromContextMenu()
        End Sub

        Private Sub SelectThreadFromContextMenu()
            Dim target = ResolveContextThreadEntry()
            If target Is Nothing Then
                Return
            End If

            SelectThreadEntry(target, suppressAutoLoad:=False)
        End Sub

        Private Async Function RefreshThreadFromContextMenuAsync() As Task
            Await _threadWorkflowCoordinator.RunThreadContextActionAsync(
                AddressOf ResolveContextThreadEntry,
                "Select a thread first.",
                Async Function(target)
                    SelectThreadEntry(target, suppressAutoLoad:=True)
                    Await AutoLoadThreadSelectionAsync(target, forceReload:=True)
                End Function)
        End Function

        Private Async Function ForkThreadFromContextMenuAsync() As Task
            Await _threadWorkflowCoordinator.RunThreadContextActionAsync(
                AddressOf ResolveContextThreadEntry,
                "Select a thread first.",
                Function(target) ForkThreadAsync(target))
        End Function

        Private Async Function ArchiveThreadFromContextMenuAsync() As Task
            Await _threadWorkflowCoordinator.RunThreadContextActionAsync(
                AddressOf ResolveContextThreadEntry,
                "Select a thread first.",
                Function(target) ArchiveThreadAsync(target))
        End Function

        Private Async Function UnarchiveThreadFromContextMenuAsync() As Task
            Await _threadWorkflowCoordinator.RunThreadContextActionAsync(
                AddressOf ResolveContextThreadEntry,
                "Select a thread first.",
                Function(target) UnarchiveThreadAsync(target))
        End Function

        Private Async Function StartThreadFromGroupHeaderContextMenuAsync() As Task
            Await _threadWorkflowCoordinator.RunThreadContextActionAsync(
                Function()
                    Dim target = ResolveContextThreadGroupEntry()
                    If target Is Nothing OrElse String.IsNullOrWhiteSpace(target.ProjectPath) Then
                        Return Nothing
                    End If

                    Return target
                End Function,
                "Select a project folder first.",
                Async Function(target)
                    ShowWorkspaceView()
                    _newThreadTargetOverrideCwd = target.ProjectPath
                    SyncNewThreadTargetChip()
                    Await StartThreadAsync()
                End Function)
        End Function

        Private Function ResolveContextThreadEntry() As ThreadListEntry
            If _threadContextTarget IsNot Nothing Then
                Return _threadContextTarget
            End If

            Return TryCast(_viewModel.ThreadsPanel.SelectedListItem, ThreadListEntry)
        End Function

        Private Function ResolveContextThreadGroupEntry() As ThreadGroupHeaderEntry
            If _threadGroupContextTarget IsNot Nothing Then
                Return _threadGroupContextTarget
            End If

            Return TryCast(_viewModel.ThreadsPanel.SelectedListItem, ThreadGroupHeaderEntry)
        End Function

        Private Sub SelectThreadEntry(entry As ThreadListEntry, suppressAutoLoad As Boolean)
            If entry Is Nothing Then
                Return
            End If

            If suppressAutoLoad Then
                _suppressThreadSelectionEvents = True
            End If

            Try
                _viewModel.ThreadsPanel.SelectedListItem = entry
            Finally
                If suppressAutoLoad Then
                    _suppressThreadSelectionEvents = False
                End If
            End Try
        End Sub

        Private Sub SetPendingNewThreadFirstPromptSelectionActive(isActive As Boolean,
                                                                  Optional clearThreadSelection As Boolean = False)
            If clearThreadSelection Then
                _suppressThreadSelectionEvents = True
                Try
                    _viewModel.ThreadsPanel.SelectedListItem = Nothing
                Finally
                    _suppressThreadSelectionEvents = False
                End Try
            End If

            If _pendingNewThreadFirstPromptSelection = isActive Then
                Return
            End If

            _pendingNewThreadFirstPromptSelection = isActive
            UpdateSidebarSelectionState(showSettings:=(_viewModel.SidebarSettingsViewVisibility = Visibility.Visible))
            UpdateThreadTurnLabels()
            UpdateWorkspaceEmptyStateVisibility()
        End Sub

        Private Function FindVisibleThreadListEntryById(threadId As String) As ThreadListEntry
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return Nothing
            End If

            For Each item As Object In _viewModel.ThreadsPanel.Items
                Dim entry = TryCast(item, ThreadListEntry)
                If entry IsNot Nothing AndAlso StringComparer.Ordinal.Equals(entry.Id, normalizedThreadId) Then
                    Return entry
                End If
            Next

            Return Nothing
        End Function

        Private Sub FinalizePendingNewThreadFirstPromptSelection()
            If Not _pendingNewThreadFirstPromptSelection Then
                Return
            End If

            SetPendingNewThreadFirstPromptSelectionActive(False)

            Dim visibleEntry = FindVisibleThreadListEntryById(_currentThreadId)
            If visibleEntry IsNot Nothing Then
                SelectThreadEntry(visibleEntry, suppressAutoLoad:=True)
            End If
        End Sub

        Private Async Function EnsurePendingDraftThreadCreatedAsync() As Task
            If Not _pendingNewThreadFirstPromptSelection OrElse Not String.IsNullOrWhiteSpace(_currentThreadId) Then
                Return
            End If

            Dim targetCwd = ResolveNewThreadTargetCwd()
            Dim options = BuildThreadRequestOptions(True)
            If Not String.IsNullOrWhiteSpace(targetCwd) Then
                options.Cwd = targetCwd
            End If

            Dim threadObject = Await _threadService.StartThreadAsync(options, CancellationToken.None).ConfigureAwait(True)
            ApplyCurrentThreadFromThreadObject(threadObject, clearPendingNewThreadSelection:=False)
        End Function

        Private Shared Function FindVisualAncestor(Of T As DependencyObject)(start As DependencyObject) As T
            Dim current = start
            While current IsNot Nothing
                Dim match = TryCast(current, T)
                If match IsNot Nothing Then
                    Return match
                End If

                current = VisualTreeHelper.GetParent(current)
            End While

            Return Nothing
        End Function

        Private Async Function ForkThreadAsync(selected As ThreadListEntry) As Task
            Await _threadWorkflowCoordinator.RunForkThreadAsync(
                selected,
                AddressOf BuildThreadRequestOptions,
                Function(threadId, options, token)
                    Return _threadService.ForkThreadAsync(threadId, options, token)
                End Function,
                AddressOf ApplyCurrentThreadFromThreadObject,
                AddressOf RenderThreadObject,
                Sub(ignoredMessage)
                    AppendSystemMessage($"Forked into new thread {_currentThreadId}.")
                End Sub,
                Sub(ignoredMessage, isError, displayToast)
                    ShowStatus($"Forked thread {_currentThreadId}.", isError:=isError, displayToast:=displayToast)
                End Sub,
                AddressOf RefreshThreadsAsync)
        End Function

        Private Async Function ArchiveThreadAsync(selected As ThreadListEntry) As Task
            Await _threadWorkflowCoordinator.RunArchiveThreadAsync(
                selected,
                Function(threadId, token)
                    Return _threadService.ArchiveThreadAsync(threadId, token)
                End Function,
                AddressOf AppendSystemMessage,
                Sub(message, isError, displayToast)
                    ShowStatus(message, isError:=isError, displayToast:=displayToast)
                End Sub,
                AddressOf RefreshThreadsAsync)
        End Function

        Private Async Function UnarchiveThreadAsync(selected As ThreadListEntry) As Task
            Await _threadWorkflowCoordinator.RunUnarchiveThreadAsync(
                selected,
                Function(threadId, token)
                    Return _threadService.UnarchiveThreadAsync(threadId, token)
                End Function,
                AddressOf AppendSystemMessage,
                Sub(message, isError, displayToast)
                    ShowStatus(message, isError:=isError, displayToast:=displayToast)
                End Sub,
                AddressOf RefreshThreadsAsync)
        End Function

        Private Sub ApplyThreadFiltersAndSort()
            Dim searchText = _viewModel.ThreadsPanel.SearchText.Trim()
            Dim forceExpandMatchingGroups = Not String.IsNullOrWhiteSpace(searchText)
            Dim filtered As New List(Of ThreadListEntry)()

            For Each entry In _threadEntries
                If MatchesThreadSearch(entry, searchText) Then
                    filtered.Add(entry)
                End If
            Next

            filtered.Sort(AddressOf CompareThreadEntries)

            Dim selectedThreadId As String = If(_pendingNewThreadFirstPromptSelection, String.Empty, _currentThreadId)
            If Not _pendingNewThreadFirstPromptSelection Then
                Dim selectedEntry = TryCast(_viewModel.ThreadsPanel.SelectedListItem, ThreadListEntry)
                If selectedEntry IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(selectedEntry.Id) Then
                    selectedThreadId = selectedEntry.Id
                End If
            End If

            Dim groupedByProject As New Dictionary(Of String, ThreadProjectGroup)(StringComparer.OrdinalIgnoreCase)
            For Each entry In filtered
                Dim groupKey = GetProjectGroupKey(entry.Cwd)
                Dim group As ThreadProjectGroup = Nothing
                If Not groupedByProject.TryGetValue(groupKey, group) Then
                    group = New ThreadProjectGroup() With {
                        .Key = groupKey,
                        .HeaderLabel = BuildProjectGroupLabel(entry.Cwd)
                    }
                    groupedByProject.Add(groupKey, group)
                End If

                group.Threads.Add(entry)
                If entry.LastActiveSortTimestamp > group.LatestActivitySortTimestamp Then
                    group.LatestActivitySortTimestamp = entry.LastActiveSortTimestamp
                End If
            Next

            Dim orderedGroups As New List(Of ThreadProjectGroup)(groupedByProject.Values)
            orderedGroups.Sort(AddressOf CompareThreadProjectGroups)

            _suppressThreadSelectionEvents = True
            Try
                _viewModel.ThreadsPanel.Items.Clear()
                For Each group In orderedGroups
                    Dim isExpanded = forceExpandMatchingGroups OrElse _expandedThreadProjectGroups.Contains(group.Key)
                    _viewModel.ThreadsPanel.Items.Add(New ThreadGroupHeaderEntry() With {
                        .GroupKey = group.Key,
                        .ProjectPath = If(StringComparer.Ordinal.Equals(group.Key, "(no-project)"), String.Empty, group.Key),
                        .FolderName = group.HeaderLabel,
                        .Count = group.Threads.Count,
                        .IsExpanded = isExpanded
                    })

                    If isExpanded Then
                        For Each entry In group.Threads
                            _viewModel.ThreadsPanel.Items.Add(entry)
                        Next
                    End If
                Next

                If Not String.IsNullOrWhiteSpace(selectedThreadId) Then
                    For i = 0 To _viewModel.ThreadsPanel.Items.Count - 1
                        Dim entry = TryCast(_viewModel.ThreadsPanel.Items(i), ThreadListEntry)
                        If entry IsNot Nothing AndAlso StringComparer.Ordinal.Equals(entry.Id, selectedThreadId) Then
                            _viewModel.ThreadsPanel.SelectedListItem = entry
                            Exit For
                        End If
                    Next
                ElseIf _pendingNewThreadFirstPromptSelection Then
                    _viewModel.ThreadsPanel.SelectedListItem = Nothing
                End If
            Finally
                _suppressThreadSelectionEvents = False
            End Try

            If TryCast(_viewModel.ThreadsPanel.SelectedListItem, ThreadListEntry) Is Nothing Then
                CancelActiveThreadSelectionLoad()
                ResetThreadSelectionLoadUiState(hideTranscriptLoader:=True)
            End If

            UpdateThreadsStateLabel(VisibleThreadCount())
            RefreshControlStates()
        End Sub

        Private Sub ToggleThreadProjectGroupExpansion(groupKey As String)
            If String.IsNullOrWhiteSpace(groupKey) Then
                Return
            End If

            If _expandedThreadProjectGroups.Contains(groupKey) Then
                _expandedThreadProjectGroups.Remove(groupKey)
            Else
                _expandedThreadProjectGroups.Add(groupKey)
            End If
        End Sub

        Private Function CompareThreadProjectGroups(left As ThreadProjectGroup, right As ThreadProjectGroup) As Integer
            If left Is Nothing AndAlso right Is Nothing Then
                Return 0
            End If

            If left Is Nothing Then
                Return -1
            End If

            If right Is Nothing Then
                Return 1
            End If

            Dim result As Integer
            Select Case _viewModel.ThreadsPanel.SortIndex
                Case 1
                    result = left.LatestActivitySortTimestamp.CompareTo(right.LatestActivitySortTimestamp)
                Case 2, 3
                    result = StringComparer.OrdinalIgnoreCase.Compare(left.HeaderLabel, right.HeaderLabel)
                Case Else
                    result = right.LatestActivitySortTimestamp.CompareTo(left.LatestActivitySortTimestamp)
            End Select

            If result <> 0 Then
                Return result
            End If

            Return StringComparer.OrdinalIgnoreCase.Compare(left.Key, right.Key)
        End Function

        Private Function CompareThreadEntries(left As ThreadListEntry, right As ThreadListEntry) As Integer
            If left Is Nothing AndAlso right Is Nothing Then
                Return 0
            End If

            If left Is Nothing Then
                Return -1
            End If

            If right Is Nothing Then
                Return 1
            End If

            Dim result As Integer
            Select Case _viewModel.ThreadsPanel.SortIndex
                Case 1
                    result = left.LastActiveSortTimestamp.CompareTo(right.LastActiveSortTimestamp)
                Case 2
                    result = StringComparer.OrdinalIgnoreCase.Compare(left.Preview, right.Preview)
                Case 3
                    result = StringComparer.OrdinalIgnoreCase.Compare(right.Preview, left.Preview)
                Case Else
                    result = right.LastActiveSortTimestamp.CompareTo(left.LastActiveSortTimestamp)
            End Select

            If result <> 0 Then
                Return result
            End If

            result = right.LastActiveSortTimestamp.CompareTo(left.LastActiveSortTimestamp)
            If result <> 0 Then
                Return result
            End If

            Return StringComparer.OrdinalIgnoreCase.Compare(left.Id, right.Id)
        End Function

        Private Function MatchesThreadSearch(entry As ThreadListEntry, searchText As String) As Boolean
            If entry Is Nothing Then
                Return False
            End If

            If String.IsNullOrWhiteSpace(searchText) Then
                Return True
            End If

            Return ContainsIgnoreCase(entry.Id, searchText) OrElse
                   ContainsIgnoreCase(entry.Preview, searchText) OrElse
                   ContainsIgnoreCase(entry.LastActiveAt, searchText) OrElse
                   ContainsIgnoreCase(entry.Cwd, searchText) OrElse
                   ContainsIgnoreCase(BuildProjectGroupLabel(entry.Cwd), searchText)
        End Function

        Private Shared Function ContainsIgnoreCase(value As String, searchText As String) As Boolean
            If String.IsNullOrWhiteSpace(value) OrElse String.IsNullOrWhiteSpace(searchText) Then
                Return False
            End If

            Return value.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Private Shared Function GetProjectGroupKey(cwd As String) As String
            Dim normalized = NormalizeProjectPath(cwd)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return "(no-project)"
            End If

            Return normalized
        End Function

        Private Shared Function BuildProjectGroupLabel(cwd As String) As String
            Dim normalized = NormalizeProjectPath(cwd)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return "No project"
            End If

            Dim folderName = Path.GetFileName(normalized)
            If String.IsNullOrWhiteSpace(folderName) Then
                Dim root = String.Empty
                Try
                    root = Path.GetPathRoot(normalized)
                Catch
                End Try

                If Not String.IsNullOrWhiteSpace(root) AndAlso
                   StringComparer.OrdinalIgnoreCase.Equals(root, normalized) Then
                    folderName = normalized.TrimEnd("\"c, "/"c)
                Else
                    folderName = normalized
                End If
            End If

            Return folderName
        End Function

        Private Shared Function NormalizeProjectPath(cwd As String) As String
            If String.IsNullOrWhiteSpace(cwd) Then
                Return String.Empty
            End If

            Dim normalized = cwd.Trim()
            Dim root = String.Empty

            Try
                root = Path.GetPathRoot(normalized)
            Catch
            End Try

            If Not String.IsNullOrWhiteSpace(root) AndAlso
               StringComparer.OrdinalIgnoreCase.Equals(normalized, root) Then
                Return normalized
            End If

            Return normalized.TrimEnd("\"c, "/"c)
        End Function

        Private Shared Function ExtractThreadWorkingDirectoryFromThreadObject(threadObject As JsonObject) As String
            If threadObject Is Nothing Then
                Return String.Empty
            End If

            Dim directKeys = {
                "cwd",
                "workingDirectory",
                "workingDir"
            }

            For Each key In directKeys
                Dim value = GetPropertyString(threadObject, key)
                If Not String.IsNullOrWhiteSpace(value) Then
                    Return NormalizeProjectPath(value)
                End If
            Next

            Dim nestedCandidates = {
                GetNestedProperty(threadObject, "context", "cwd"),
                GetNestedProperty(threadObject, "context", "workingDirectory"),
                GetNestedProperty(threadObject, "workspace", "cwd"),
                GetNestedProperty(threadObject, "workspace", "workingDirectory"),
                GetNestedProperty(threadObject, "project", "cwd"),
                GetNestedProperty(threadObject, "project", "workingDirectory")
            }

            For Each candidate In nestedCandidates
                Dim value As String = String.Empty
                If TryGetStringValue(candidate, value) AndAlso Not String.IsNullOrWhiteSpace(value) Then
                    Return NormalizeProjectPath(value)
                End If
            Next

            Return String.Empty
        End Function

        Private Function VisibleThreadCount() As Integer
            Dim count = 0
            For Each item As Object In _viewModel.ThreadsPanel.Items
                If TypeOf item Is ThreadListEntry Then
                    count += 1
                End If
            Next

            Return count
        End Function

        Private Sub UpdateThreadsStateLabel(displayCount As Integer)
            SyncSessionStateViewModel()
            Dim session = _viewModel.SessionState
            Dim connected = session.IsConnected

            If Not connected Then
                _viewModel.ThreadsPanel.StateText = "Connect to Codex App Server to load threads."
                Return
            End If

            Dim hasProjectHeaders = False
            If displayCount = 0 Then
                For Each item As Object In _viewModel.ThreadsPanel.Items
                    If TypeOf item Is ThreadGroupHeaderEntry Then
                        hasProjectHeaders = True
                        Exit For
                    End If
                Next
            End If

            _viewModel.ThreadsPanel.UpdateThreadListStateText(connected,
                                                              session.IsAuthenticated,
                                                              _threadsLoading,
                                                              _viewModel.ThreadsPanel.RefreshErrorText,
                                                              _threadEntries.Count,
                                                              displayCount,
                                                              hasProjectHeaders)
        End Sub

        Private Sub MarkThreadLastActive(threadId As String, Optional unixMilliseconds As Long = 0)
            If String.IsNullOrWhiteSpace(threadId) Then
                Return
            End If

            If unixMilliseconds <= 0 Then
                unixMilliseconds = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()
            End If

            Dim localTimestamp As String
            Try
                localTimestamp = DateTimeOffset.FromUnixTimeMilliseconds(unixMilliseconds).LocalDateTime.ToString("yyyy-MM-dd HH:mm:ss")
            Catch
                localTimestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            End Try

            Dim changed = False
            For Each entry In _threadEntries
                If StringComparer.Ordinal.Equals(entry.Id, threadId) Then
                    entry.LastActiveSortTimestamp = unixMilliseconds
                    entry.LastActiveAt = localTimestamp
                    changed = True
                    Exit For
                End If
            Next

            If changed Then
                ApplyThreadFiltersAndSort()
            End If
        End Sub

        Private Function SyncThreadListAfterUserPrompt(threadId As String, promptText As String) As Boolean
            If String.IsNullOrWhiteSpace(threadId) Then
                Return False
            End If

            Dim normalizedPrompt = NormalizeThreadPreviewFromPrompt(promptText)
            Dim foundEntry = False
            Dim changed = False

            For Each entry In _threadEntries
                If Not StringComparer.Ordinal.Equals(entry.Id, threadId) Then
                    Continue For
                End If

                foundEntry = True

                If Not String.IsNullOrWhiteSpace(normalizedPrompt) AndAlso
                   String.IsNullOrWhiteSpace(entry.Preview) Then
                    entry.Preview = normalizedPrompt
                    changed = True
                End If

                Exit For
            Next

            If changed Then
                ApplyThreadFiltersAndSort()
            End If

            ' If the thread is not yet present, the server likely didn't list it until after the first turn.
            Return Not foundEntry
        End Function

        Private Shared Function NormalizeThreadPreviewFromPrompt(promptText As String) As String
            If String.IsNullOrWhiteSpace(promptText) Then
                Return String.Empty
            End If

            Dim normalized = promptText.Replace(ControlChars.Cr, " "c).
                                        Replace(ControlChars.Lf, " "c).
                                        Replace(ControlChars.Tab, " "c).
                                        Trim()

            Do While normalized.Contains("  ", StringComparison.Ordinal)
                normalized = normalized.Replace("  ", " ", StringComparison.Ordinal)
            Loop

            Return normalized
        End Function

        Private Function BuildThreadRequestOptions(includeModel As Boolean) As ThreadRequestOptions
            Dim options As New ThreadRequestOptions() With {
                .ApprovalPolicy = _viewModel.TurnComposer.SelectedApprovalPolicy,
                .Sandbox = _viewModel.TurnComposer.SelectedSandbox,
                .Cwd = EffectiveThreadWorkingDirectory()
            }

            If includeModel Then
                options.Model = _viewModel.TurnComposer.SelectedModelId
            End If

            Return options
        End Function

        Private Sub RenderThreadObject(threadObject As JsonObject)
            _viewModel.TranscriptPanel.ClearTranscript()

            Dim turns = GetPropertyArray(threadObject, "turns")
            If turns Is Nothing OrElse turns.Count = 0 Then
                AppendSystemMessage("No historical turns loaded for this thread.")
                Return
            End If

            For Each turnNode In turns
                Dim turnObject = AsObject(turnNode)
                If turnObject Is Nothing Then
                    Continue For
                End If

                Dim items = GetPropertyArray(turnObject, "items")
                If items Is Nothing Then
                    Continue For
                End If

                For Each itemNode In items
                    Dim itemObject = AsObject(itemNode)
                    If itemObject IsNot Nothing Then
                        RenderItem(itemObject)
                    End If
                Next
            Next

            ScrollTranscriptToBottom()
        End Sub

        Private Sub ApplyThreadTranscriptSnapshot(transcriptSnapshot As ThreadTranscriptSnapshotData, hasTurns As Boolean)
            _viewModel.TranscriptPanel.ClearTranscript()
            If Not hasTurns Then
                AppendSystemMessage("No historical turns loaded for this thread.")
                Return
            End If

            Dim snapshot = If(transcriptSnapshot, New ThreadTranscriptSnapshotData())
            _viewModel.TranscriptPanel.SetTranscriptSnapshot(snapshot.RawText)
            _viewModel.TranscriptPanel.SetTranscriptDisplaySnapshot(snapshot.DisplayEntries)
            ScrollTranscriptToBottom()
        End Sub

        Private Shared Function ThreadObjectHasTurns(threadObject As JsonObject) As Boolean
            Dim turns = GetPropertyArray(threadObject, "turns")
            Return turns IsNot Nothing AndAlso turns.Count > 0
        End Function

        Private Shared Function BuildThreadTranscriptSnapshot(threadObject As JsonObject) As ThreadTranscriptSnapshotData
            Dim snapshot As New ThreadTranscriptSnapshotData()
            Dim turns = GetPropertyArray(threadObject, "turns")
            If turns Is Nothing OrElse turns.Count = 0 Then
                Return snapshot
            End If

            Dim builder As New StringBuilder()
            For Each turnNode In turns
                Dim turnObject = AsObject(turnNode)
                If turnObject Is Nothing Then
                    Continue For
                End If

                Dim items = GetPropertyArray(turnObject, "items")
                If items Is Nothing Then
                    Continue For
                End If

                For Each itemNode In items
                    Dim itemObject = AsObject(itemNode)
                    If itemObject IsNot Nothing Then
                        AppendSnapshotItem(builder, snapshot.DisplayEntries, itemObject)
                    End If
                Next
            Next

            snapshot.RawText = builder.ToString().TrimEnd()
            Return snapshot
        End Function

        Private Shared Sub AppendSnapshotItem(builder As StringBuilder,
                                              entries As IList(Of TranscriptEntryDescriptor),
                                              itemObject As JsonObject)
            Dim itemType = GetPropertyString(itemObject, "type")

            Select Case itemType
                Case "userMessage"
                    Dim content = GetPropertyArray(itemObject, "content")
                    Dim text = FlattenUserInput(content)
                    AppendSnapshotLine(builder, "user", text)
                    AddSnapshotDescriptor(entries, "user", "You", text)

                Case "agentMessage"
                    Dim text = GetPropertyString(itemObject, "text")
                    AppendSnapshotLine(builder, "assistant", text)
                    If IsCommentaryAgentMessage(itemObject) Then
                        If entries IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(text) Then
                            entries.Add(New TranscriptEntryDescriptor() With {
                                .Kind = "reasoning",
                                .RoleText = "Reasoning",
                                .BodyText = text,
                                .DetailsText = String.Empty,
                                .IsMuted = True,
                                .IsReasoning = True
                            })
                        End If
                    Else
                        AddSnapshotDescriptor(entries, "assistant", "Codex", text)
                    End If

                Case "plan"
                    Dim text = GetPropertyString(itemObject, "text")
                    AppendSnapshotLine(builder, "plan", text)
                    AddSnapshotDescriptor(entries, "plan", "Plan", text)

                Case "reasoning"
                    Dim text = ExtractReasoningText(itemObject)
                    If Not String.IsNullOrWhiteSpace(text) Then
                        AppendSnapshotLine(builder, "reasoning", text)
                        If entries IsNot Nothing Then
                            entries.Add(New TranscriptEntryDescriptor() With {
                                .Kind = "reasoning",
                                .RoleText = "Reasoning",
                                .BodyText = text,
                                .DetailsText = text,
                                .IsMuted = True,
                                .IsReasoning = True
                            })
                        End If
                    End If

                Case "commandExecution"
                    Dim command = GetPropertyString(itemObject, "command")
                    Dim status = GetPropertyString(itemObject, "status")
                    Dim output = GetPropertyString(itemObject, "aggregatedOutput")
                    Dim summary As New StringBuilder()
                    summary.AppendLine($"Command ({status}): {command}")
                    If Not String.IsNullOrWhiteSpace(output) Then
                        summary.AppendLine(output)
                    End If
                    AppendSnapshotLine(builder, "command", summary.ToString().TrimEnd())
                    If entries IsNot Nothing Then
                        entries.Add(New TranscriptEntryDescriptor() With {
                            .Kind = "command",
                            .RoleText = "Command",
                            .BodyText = If(String.IsNullOrWhiteSpace(command), "(command)", command),
                            .SecondaryText = If(String.IsNullOrWhiteSpace(status), String.Empty, $"status: {status}"),
                            .DetailsText = If(output, String.Empty).Trim(),
                            .IsCommandLike = True
                        })
                    End If

                Case "fileChange"
                    Dim status = GetPropertyString(itemObject, "status")
                    Dim changes = GetPropertyArray(itemObject, "changes")
                    Dim count = If(changes Is Nothing, 0, changes.Count)
                    Dim lineStats = BuildFileChangeLineStats(changes)
                    AppendSnapshotLine(builder, "fileChange", $"{count} change(s), status={status}")
                    If entries IsNot Nothing Then
                        entries.Add(New TranscriptEntryDescriptor() With {
                            .Kind = "fileChange",
                            .RoleText = "Files",
                            .BodyText = $"{count} change(s){If(String.IsNullOrWhiteSpace(status), String.Empty, $" ({status})")}",
                            .DetailsText = BuildSnapshotFileChangeDetails(changes),
                            .AddedLineCount = lineStats.AddedLineCount,
                            .RemovedLineCount = lineStats.RemovedLineCount
                        })
                    End If

                Case Else
                    Dim itemId = GetPropertyString(itemObject, "id")
                    AppendSnapshotLine(builder, "item", $"{itemType} ({itemId})")
                    If entries IsNot Nothing Then
                        entries.Add(New TranscriptEntryDescriptor() With {
                            .Kind = "event",
                            .RoleText = "Item",
                            .BodyText = $"{itemType} ({itemId})",
                            .IsMuted = True
                        })
                    End If
            End Select
        End Sub

        Private Shared Sub AddSnapshotDescriptor(entries As IList(Of TranscriptEntryDescriptor),
                                                 kind As String,
                                                 roleText As String,
                                                 bodyText As String)
            If entries Is Nothing OrElse String.IsNullOrWhiteSpace(bodyText) Then
                Return
            End If

            entries.Add(New TranscriptEntryDescriptor() With {
                .Kind = If(kind, String.Empty),
                .RoleText = If(roleText, String.Empty),
                .BodyText = bodyText
            })
        End Sub

        Private Shared Function BuildSnapshotFileChangeDetails(changes As JsonArray) As String
            If changes Is Nothing OrElse changes.Count = 0 Then
                Return String.Empty
            End If

            Dim builder As New StringBuilder()
            Dim shown = 0
            For Each changeNode In changes
                If shown >= 12 Then
                    Exit For
                End If

                Dim changeObject = AsObject(changeNode)
                If changeObject Is Nothing Then
                    Continue For
                End If

                Dim path = GetPropertyString(changeObject, "path")
                If String.IsNullOrWhiteSpace(path) Then
                    path = GetPropertyString(changeObject, "file")
                End If
                If String.IsNullOrWhiteSpace(path) Then
                    Continue For
                End If

                Dim status = GetPropertyString(changeObject, "status")
                If String.IsNullOrWhiteSpace(status) Then
                    builder.AppendLine(path)
                Else
                    builder.AppendLine($"{status}: {path}")
                End If

                shown += 1
            Next

            If builder.Length = 0 Then
                Return String.Empty
            End If

            If changes.Count > shown Then
                builder.Append("... +")
                builder.Append(changes.Count - shown)
                builder.Append(" more")
            End If

            Return builder.ToString().TrimEnd()
        End Function

        Private Shared Sub AppendSnapshotLine(builder As StringBuilder, role As String, text As String)
            If String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            builder.Append("["c)
            builder.Append(Date.Now.ToString("HH:mm:ss"))
            builder.Append("] ")
            builder.Append(role)
            builder.Append(": ")
            builder.Append(text)
            builder.AppendLine()
            builder.AppendLine()
        End Sub

        Private Sub ApplyCurrentThreadFromThreadObject(threadObject As JsonObject,
                                                       Optional clearPendingNewThreadSelection As Boolean = True)
            If clearPendingNewThreadSelection Then
                SetPendingNewThreadFirstPromptSelectionActive(False)
            End If

            Dim threadId = GetPropertyString(threadObject, "id")
            If Not String.IsNullOrWhiteSpace(threadId) Then
                _currentThreadId = threadId
            End If

            Dim loadedThreadCwd = ExtractThreadWorkingDirectoryFromThreadObject(threadObject)
            If Not String.IsNullOrWhiteSpace(loadedThreadCwd) Then
                _currentThreadCwd = loadedThreadCwd
                _newThreadTargetOverrideCwd = String.Empty
            ElseIf String.IsNullOrWhiteSpace(_newThreadTargetOverrideCwd) Then
                _currentThreadCwd = String.Empty
            End If

            _currentTurnId = String.Empty
            UpdateThreadTurnLabels()
            RefreshControlStates()
        End Sub

        Private Sub EnsureThreadSelected()
            If String.IsNullOrWhiteSpace(_currentThreadId) Then
                Throw New InvalidOperationException("No active thread selected.")
            End If
        End Sub

        Private Function SelectedThreadEntry() As ThreadListEntry
            Dim selected = TryCast(_viewModel.ThreadsPanel.SelectedListItem, ThreadListEntry)
            If selected Is Nothing OrElse String.IsNullOrWhiteSpace(selected.Id) Then
                Throw New InvalidOperationException("Select a thread first.")
            End If

            Return selected
        End Function

    End Class
End Namespace
