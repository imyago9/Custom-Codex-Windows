Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
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
            Public Property TranscriptSnapshot As ThreadTranscriptSnapshot
        End Class

        Private Function StartThreadAsync() As Task
            CancelActiveThreadSelectionLoad()
            ResetThreadSelectionLoadUiState(hideTranscriptLoader:=True)
            ClearPendingUserEchoTracking()
            _viewModel.TranscriptPanel.ClearTranscript()
            ClearVisibleThreadId()
            ClearVisibleTurnId()
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

        Private Async Function ResumeThreadForSelectionAsync(threadId As String,
                                                             cancellationToken As CancellationToken) As Task(Of JsonObject)
            Return Await _threadService.ResumeThreadAsync(threadId,
                                                          New ThreadRequestOptions(),
                                                          cancellationToken).ConfigureAwait(False)
        End Function

        Private Async Function ReadThreadForSelectionAsync(threadId As String,
                                                           cancellationToken As CancellationToken) As Task(Of JsonObject)
            Return Await _threadService.ReadThreadAsync(threadId,
                                                        includeTurns:=True,
                                                        cancellationToken:=cancellationToken).ConfigureAwait(False)
        End Function

        Private Async Function LoadThreadObjectForSelectionAsync(threadId As String,
                                                                 cancellationToken As CancellationToken) As Task(Of JsonObject)
            Dim resumedThread = Await ResumeThreadForSelectionAsync(threadId, cancellationToken).ConfigureAwait(False)
            If ThreadObjectHasTurns(resumedThread) Then
                Return resumedThread
            End If

            Return Await ReadThreadForSelectionAsync(threadId, cancellationToken).ConfigureAwait(False)
        End Function

        Private Async Function LoadThreadSelectionPayloadAsync(threadId As String,
                                                               cancellationToken As CancellationToken) As Task(Of ThreadSelectionLoadPayload)
            ' Build the historical snapshot from thread/read before resume loads the thread and may lock the rollout JSONL.
            Dim snapshotSourceThread = Await ReadThreadForSelectionAsync(threadId, cancellationToken).ConfigureAwait(False)
            cancellationToken.ThrowIfCancellationRequested()

            Dim hasTurns = ThreadObjectHasTurns(snapshotSourceThread)
            Dim transcriptSnapshot As ThreadTranscriptSnapshot = Nothing
            If hasTurns Then
                transcriptSnapshot = Await Task.Run(Function() BuildThreadTranscriptSnapshot(snapshotSourceThread), cancellationToken).ConfigureAwait(False)
                cancellationToken.ThrowIfCancellationRequested()
            Else
                transcriptSnapshot = New ThreadTranscriptSnapshot()
            End If

            Dim threadObject = Await ResumeThreadForSelectionAsync(threadId, cancellationToken).ConfigureAwait(False)
            cancellationToken.ThrowIfCancellationRequested()

            If threadObject Is Nothing Then
                threadObject = snapshotSourceThread
            End If

            If Not ThreadObjectHasTurns(threadObject) AndAlso hasTurns Then
                ' Keep the snapshot source thread object as a fallback if resume omits turns.
                threadObject = snapshotSourceThread
            End If

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
                    PersistThreadSelectionSnapshotToLiveRegistry(request.ThreadId, payload.TranscriptSnapshot)
                    RebuildVisibleTranscriptForThread(request.ThreadId)
                    If Not payload.HasTurns AndAlso _viewModel.TranscriptPanel.Items.Count = 0 Then
                        AppendSystemMessage("No historical turns loaded for this thread.")
                    End If
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
                _suppressThreadSelectionEvents = True
                Try
                    _viewModel.ThreadsPanel.SelectedListItem = header
                Finally
                    _suppressThreadSelectionEvents = False
                End Try
                e.Handled = True
                PrepareThreadGroupContextMenu(header)
                SidebarPaneHost.ThreadItemContextMenu.PlacementTarget = container
                SidebarPaneHost.ThreadItemContextMenu.IsOpen = True
                Return
            End If

            Dim entry = TryCast(container.DataContext, ThreadListEntry)
            If entry IsNot Nothing Then
                _threadContextTarget = entry
                SelectThreadEntry(entry, suppressAutoLoad:=True)
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
                    Return ResolveContextThreadGroupEntryWithProjectPath()
                End Function,
                "Select a project folder first.",
                Async Function(target)
                    ShowWorkspaceView()
                    _newThreadTargetOverrideCwd = target.ProjectPath
                    SyncNewThreadTargetChip()
                    Await StartThreadAsync()
                End Function)
        End Function

        Private Async Function OpenThreadGroupInVsCodeFromContextMenuAsync() As Task
            Await _threadWorkflowCoordinator.RunThreadContextActionAsync(
                Function()
                    Return ResolveContextThreadGroupEntryWithProjectPath()
                End Function,
                "Select a project folder first.",
                Function(target)
                    If Not Directory.Exists(target.ProjectPath) Then
                        ShowStatus($"Folder not found: {target.ProjectPath}", isError:=True, displayToast:=True)
                        Return Task.CompletedTask
                    End If

                    If StartVsCode(target.ProjectPath, ".") Then
                        ShowStatus($"Opened VS Code in {target.ProjectPath}")
                        Return Task.CompletedTask
                    End If

                    ShowStatus("Could not open VS Code. Make sure `code` is installed and available on PATH.",
                               isError:=True,
                               displayToast:=True)
                    Return Task.CompletedTask
                End Function)
        End Function

        Private Async Function OpenThreadGroupInTerminalFromContextMenuAsync() As Task
            Await _threadWorkflowCoordinator.RunThreadContextActionAsync(
                Function()
                    Return ResolveContextThreadGroupEntryWithProjectPath()
                End Function,
                "Select a project folder first.",
                Function(target)
                    If Not Directory.Exists(target.ProjectPath) Then
                        ShowStatus($"Folder not found: {target.ProjectPath}", isError:=True, displayToast:=True)
                        Return Task.CompletedTask
                    End If

                    If StartProcessInDirectory("powershell.exe", "-NoExit", target.ProjectPath) Then
                        ShowStatus($"Opened PowerShell in {target.ProjectPath}")
                        Return Task.CompletedTask
                    End If

                    ShowStatus("Could not open PowerShell.", isError:=True, displayToast:=True)
                    Return Task.CompletedTask
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

        Private Function ResolveContextThreadGroupEntryWithProjectPath() As ThreadGroupHeaderEntry
            Dim target = ResolveContextThreadGroupEntry()
            If target Is Nothing OrElse String.IsNullOrWhiteSpace(target.ProjectPath) Then
                Return Nothing
            End If

            Return target
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
            ' Seed an empty historical baseline for brand-new threads before the first turn starts.
            ' This prevents the first re-open during an in-flight turn from using a partial snapshot
            ' as the baseline and duplicating prompt/output when overlay replay runs.
            PersistThreadSelectionSnapshotToLiveRegistry(_currentThreadId, New ThreadTranscriptSnapshot())
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
            UpdateThreadEntryRuntimeIndicatorsFromSessionState()

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

        Private Sub RefreshThreadRuntimeIndicatorsIfNeeded()
            If _threadEntries.Count = 0 Then
                Return
            End If

            If UpdateThreadEntryRuntimeIndicatorsFromSessionState() Then
                ApplyThreadFiltersAndSort()
            End If
        End Sub

        Private Function UpdateThreadEntryRuntimeIndicatorsFromSessionState() As Boolean
            If _threadEntries.Count = 0 Then
                Return False
            End If

            Dim runtimeStore = If(_sessionNotificationCoordinator Is Nothing,
                                  Nothing,
                                  _sessionNotificationCoordinator.RuntimeStore)
            Dim visibleThreadId = If(_currentThreadId, String.Empty).Trim()
            Dim changed = False

            For Each entry In _threadEntries
                If entry Is Nothing OrElse String.IsNullOrWhiteSpace(entry.Id) Then
                    Continue For
                End If

                Dim threadId = entry.Id.Trim()
                Dim hasActiveRuntimeTurn = runtimeStore IsNot Nothing AndAlso runtimeStore.HasActiveTurn(threadId)

                Dim hasPendingRuntimeUpdates = False
                Dim liveState As ThreadLiveSessionState = Nothing
                If _threadLiveSessionRegistry.TryGet(threadId, liveState) AndAlso liveState IsNot Nothing Then
                    hasPendingRuntimeUpdates = liveState.PendingRebuild
                    If runtimeStore Is Nothing AndAlso Not hasActiveRuntimeTurn Then
                        hasActiveRuntimeTurn = liveState.IsTurnActive AndAlso
                                               Not String.IsNullOrWhiteSpace(If(liveState.ActiveTurnId, String.Empty).Trim())
                    End If
                End If

                If Not String.IsNullOrWhiteSpace(visibleThreadId) AndAlso
                   StringComparer.Ordinal.Equals(threadId, visibleThreadId) Then
                    ' "Dirty" means hidden-thread updates pending a rebind; hide it for the visible thread.
                    hasPendingRuntimeUpdates = False
                End If

                If entry.HasActiveRuntimeTurn <> hasActiveRuntimeTurn Then
                    entry.HasActiveRuntimeTurn = hasActiveRuntimeTurn
                    changed = True
                End If

                If entry.HasPendingRuntimeUpdates <> hasPendingRuntimeUpdates Then
                    entry.HasPendingRuntimeUpdates = hasPendingRuntimeUpdates
                    changed = True
                End If
            Next

            Return changed
        End Function

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

        Private Function HasThreadEntry(threadId As String) As Boolean
            If String.IsNullOrWhiteSpace(threadId) Then
                Return False
            End If

            Dim normalizedThreadId = threadId.Trim()
            For Each entry In _threadEntries
                If entry IsNot Nothing AndAlso StringComparer.Ordinal.Equals(entry.Id, normalizedThreadId) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Async Function RetrySilentThreadRefreshUntilListedAsync(threadId As String) As Task
            If String.IsNullOrWhiteSpace(threadId) OrElse HasThreadEntry(threadId) Then
                Return
            End If

            Const maxAttempts As Integer = 2
            Const retryDelayMs As Integer = 550

            For attempt = 1 To maxAttempts
                Await Task.Delay(retryDelayMs).ConfigureAwait(True)

                If HasThreadEntry(threadId) Then
                    Exit For
                End If

                Await RefreshThreadsCoreAsync(silent:=True).ConfigureAwait(True)
                If HasThreadEntry(threadId) Then
                    Exit For
                End If
            Next
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
            Dim hasTurns = ThreadObjectHasTurns(threadObject)
            Dim snapshot = BuildThreadTranscriptSnapshot(threadObject)
            ApplyThreadTranscriptSnapshot(snapshot, hasTurns)
        End Sub

        Private Sub ApplyThreadTranscriptSnapshot(transcriptSnapshot As ThreadTranscriptSnapshot, hasTurns As Boolean)
            ClearPendingUserEchoTracking()
            _viewModel.TranscriptPanel.ClearTranscript()
            If Not hasTurns Then
                AppendSystemMessage("No historical turns loaded for this thread.")
                Return
            End If

            Dim snapshot = If(transcriptSnapshot, New ThreadTranscriptSnapshot())
            MergeSnapshotAssistantPhaseHints(snapshot.AssistantPhaseHintsByItemKey)
            If snapshot.DebugMessages.Count > 0 Then
                For Each message In snapshot.DebugMessages
                    If String.IsNullOrWhiteSpace(message) Then
                        Continue For
                    End If

                    AppendProtocol("debug", message)
                Next
            End If
            _viewModel.TranscriptPanel.SetTranscriptSnapshot(snapshot.RawText)
            _viewModel.TranscriptPanel.SetTranscriptDisplaySnapshot(snapshot.DisplayEntries)
            ScrollTranscriptToBottom()
        End Sub

        Private Sub PersistThreadSelectionSnapshotToLiveRegistry(threadId As String, transcriptSnapshot As ThreadTranscriptSnapshot)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            Dim snapshot = If(transcriptSnapshot, New ThreadTranscriptSnapshot())
            MergeSnapshotAssistantPhaseHints(snapshot.AssistantPhaseHintsByItemKey)

            Dim runtimeStore = _sessionNotificationCoordinator.RuntimeStore
            If runtimeStore IsNot Nothing AndAlso runtimeStore.HasActiveTurn(normalizedThreadId) Then
                Dim existingState As ThreadLiveSessionState = Nothing
                If _threadLiveSessionRegistry.TryGet(normalizedThreadId, existingState) AndAlso
                   existingState IsNot Nothing AndAlso
                   existingState.HasLoadedHistoricalSnapshot Then
                    ' Preserve the last stable historical snapshot while the active turn is in-flight.
                    ' Rebuild will layer the live runtime overlay on top of this baseline.
                    Return
                End If
            End If

            _threadLiveSessionRegistry.UpsertSnapshot(normalizedThreadId,
                                                      snapshot.RawText,
                                                      snapshot.DisplayEntries,
                                                      GetVisibleTurnId())
        End Sub

        Private Sub RebuildVisibleTranscriptForThread(threadId As String)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            Dim projection = _threadLiveSessionRegistry.GetProjectionSnapshot(normalizedThreadId)
            ApplyOverlayRuntimeOrderMetadataToProjection(normalizedThreadId, projection, _sessionNotificationCoordinator.RuntimeStore)
            Dim completedOverlayTurnsForFullReplay = RemoveProjectionSnapshotItemRowsForCompletedOverlayTurns(normalizedThreadId,
                                                                                                              projection,
                                                                                                              _sessionNotificationCoordinator.RuntimeStore)
            RemoveProjectionOverlayReplaySupersededMarkers(normalizedThreadId,
                                                           projection,
                                                           _sessionNotificationCoordinator.RuntimeStore)

            ClearPendingUserEchoTracking()
            _viewModel.TranscriptPanel.ClearTranscript()
            _viewModel.TranscriptPanel.SetTranscriptSnapshot(projection.RawText)
            _viewModel.TranscriptPanel.SetTranscriptDisplaySnapshot(projection.DisplayEntries)

            ApplyLiveRuntimeOverlayForThread(normalizedThreadId,
                                            _sessionNotificationCoordinator.RuntimeStore,
                                            completedOverlayTurnsForFullReplay)

            _threadLiveSessionRegistry.MarkBound(normalizedThreadId, GetVisibleTurnId())
            _threadLiveSessionRegistry.SetPendingRebuild(normalizedThreadId, False)
            RefreshThreadRuntimeIndicatorsIfNeeded()
            ScrollTranscriptToBottom()
        End Sub

        Private Sub ApplyOverlayRuntimeOrderMetadataToProjection(threadId As String,
                                                                 projection As ThreadTranscriptProjectionSnapshot,
                                                                 runtimeStore As TurnFlowRuntimeStore)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse projection Is Nothing OrElse runtimeStore Is Nothing Then
                Return
            End If

            Dim overlayTurnIds = ResolveOverlayTurnIdsForReplay(normalizedThreadId, runtimeStore)
            If overlayTurnIds.Count = 0 Then
                Return
            End If

            Dim runtimeItemsByKey As New Dictionary(Of String, TurnItemRuntimeState)(StringComparer.Ordinal)
            For Each turnId In overlayTurnIds
                Dim normalizedTurnId = If(turnId, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                    Continue For
                End If

                Dim orderedTurnItems = runtimeStore.GetOrderedRuntimeItemsForTurn(normalizedThreadId, normalizedTurnId)
                AssignTurnItemOrderIndexes(orderedTurnItems)

                For Each item In orderedTurnItems
                    If item Is Nothing OrElse item.IsScopeInferred Then
                        Continue For
                    End If

                    Dim scopedItemKey = If(item.ScopedItemKey, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(scopedItemKey) Then
                        Continue For
                    End If

                    runtimeItemsByKey($"item:{scopedItemKey}") = item
                Next
            Next

            If runtimeItemsByKey.Count = 0 OrElse projection.DisplayEntries.Count = 0 Then
                Return
            End If

            For Each descriptor In projection.DisplayEntries
                If descriptor Is Nothing Then
                    Continue For
                End If

                Dim runtimeKey = If(descriptor.RuntimeKey, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(runtimeKey) Then
                    Continue For
                End If

                Dim runtimeItem As TurnItemRuntimeState = Nothing
                If Not runtimeItemsByKey.TryGetValue(runtimeKey, runtimeItem) OrElse runtimeItem Is Nothing Then
                    Continue For
                End If

                descriptor.ThreadId = If(String.IsNullOrWhiteSpace(descriptor.ThreadId),
                                         If(runtimeItem.ThreadId, String.Empty).Trim(),
                                         descriptor.ThreadId)
                descriptor.TurnId = If(String.IsNullOrWhiteSpace(descriptor.TurnId),
                                       If(runtimeItem.TurnId, String.Empty).Trim(),
                                       descriptor.TurnId)
                descriptor.TurnItemStreamSequence = runtimeItem.TurnItemStreamSequence
                descriptor.TurnItemOrderIndex = runtimeItem.TurnItemOrderIndex
                descriptor.TurnItemSortTimestampUtc = If(runtimeItem.StartedAt, runtimeItem.CompletedAt)
            Next
        End Sub

        Private Sub ApplyLiveRuntimeOverlayForThread(threadId As String,
                                                     runtimeStore As TurnFlowRuntimeStore,
                                                     Optional completedTurnsForFullItemReplay As ISet(Of String) = Nothing)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse runtimeStore Is Nothing Then
                Return
            End If

            Dim overlayTurnIds = ResolveOverlayTurnIdsForReplay(normalizedThreadId, runtimeStore)
            If overlayTurnIds.Count = 0 Then
                ' No tracked live overlay for this thread; avoid replaying stale/inferred runtime rows
                ' that can overwrite historical snapshot bubbles during thread re-open.
                UpdateTokenUsageWidget(normalizedThreadId, String.Empty, Nothing)
                Return
            End If

            Dim renderedScopedItemKeys As New HashSet(Of String)(StringComparer.Ordinal)

            For Each turn In runtimeStore.GetOrderedTurnStatesForThread(normalizedThreadId)
                If turn Is Nothing Then
                    Continue For
                End If

                Dim turnThreadId = If(turn.ThreadId, normalizedThreadId).Trim()
                Dim turnId = If(turn.TurnId, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(turnThreadId) OrElse String.IsNullOrWhiteSpace(turnId) Then
                    Continue For
                End If

                If overlayTurnIds.Count > 0 AndAlso Not overlayTurnIds.Contains(turnId) Then
                    Continue For
                End If

                Dim replayCompletedTurnFully = turn.IsCompleted AndAlso
                                              completedTurnsForFullItemReplay IsNot Nothing AndAlso
                                              completedTurnsForFullItemReplay.Contains(turnId)
                Dim replayCompletedTurnRuntimeOnly = turn.IsCompleted AndAlso Not replayCompletedTurnFully
                Dim orderedTurnItems As IReadOnlyList(Of TurnItemRuntimeState) = runtimeStore.GetOrderedRuntimeItemsForTurn(turnThreadId, turnId)
                AssignTurnItemOrderIndexes(orderedTurnItems)
                If replayCompletedTurnRuntimeOnly Then
                    orderedTurnItems = OrderCompletedTurnRuntimeOnlyReplayItems(orderedTurnItems)
                End If

                If Not replayCompletedTurnRuntimeOnly Then
                    ' Rebuild ordering should match live rendering: user prompt first, then turn started.
                    For Each item In orderedTurnItems
                        If item Is Nothing Then
                            Continue For
                        End If

                        If item.IsScopeInferred Then
                            Continue For
                        End If

                        If Not StringComparer.OrdinalIgnoreCase.Equals(If(item.ItemType, String.Empty), "userMessage") Then
                            Continue For
                        End If

                        Dim scopedKey = If(item.ScopedItemKey, String.Empty).Trim()
                        If String.IsNullOrWhiteSpace(scopedKey) OrElse Not renderedScopedItemKeys.Add(scopedKey) Then
                            Continue For
                        End If
                        RenderItem(item)
                    Next

                    AppendTurnLifecycleMarker(turnThreadId, turnId, "started")
                End If

                If Not String.IsNullOrWhiteSpace(turn.PlanSummary) Then
                    UpsertTurnMetadata(turnThreadId, turnId, "plan", turn.PlanSummary)
                End If

                If Not String.IsNullOrWhiteSpace(turn.DiffSummary) Then
                    UpsertTurnMetadata(turnThreadId, turnId, "diff", turn.DiffSummary)
                End If

                For Each item In orderedTurnItems
                    If item Is Nothing Then
                        Continue For
                    End If

                    If item.IsScopeInferred Then
                        Continue For
                    End If

                    If replayCompletedTurnRuntimeOnly AndAlso
                       ShouldSkipCompletedTurnSnapshotDuplicatedOverlayItem(item) Then
                        Continue For
                    End If

                    If StringComparer.OrdinalIgnoreCase.Equals(If(item.ItemType, String.Empty), "userMessage") Then
                        Continue For
                    End If

                    Dim scopedKey = If(item.ScopedItemKey, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(scopedKey) OrElse Not renderedScopedItemKeys.Add(scopedKey) Then
                        Continue For
                    End If
                    RenderItem(item)
                Next

                If turn.IsCompleted Then
                    Dim completionStatus = If(String.IsNullOrWhiteSpace(turn.TurnStatus), "completed", turn.TurnStatus)
                    AppendTurnLifecycleMarker(turnThreadId, turnId, completionStatus)
                End If

            Next

            ' Fallback for items that may exist without a materialized turn state yet.
            For Each item In EnumerateRuntimeItemsForThread(normalizedThreadId)
                If item Is Nothing Then
                    Continue For
                End If

                If item.IsScopeInferred Then
                    Continue For
                End If

                If overlayTurnIds.Count > 0 AndAlso
                   Not overlayTurnIds.Contains(If(item.TurnId, String.Empty).Trim()) Then
                    Continue For
                End If

                Dim itemTurnState = runtimeStore.GetTurnState(normalizedThreadId, item.TurnId)
                If itemTurnState IsNot Nothing AndAlso itemTurnState.IsCompleted AndAlso
                   ShouldSkipCompletedTurnSnapshotDuplicatedOverlayItem(item) Then
                    Continue For
                End If

                Dim scopedKey = If(item.ScopedItemKey, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(scopedKey) OrElse Not renderedScopedItemKeys.Add(scopedKey) Then
                    Continue For
                End If
                RenderItem(item)
            Next

            If overlayTurnIds.Count = 0 Then
                UpdateTokenUsageWidget(normalizedThreadId, String.Empty, Nothing)
                Return
            End If

            Dim threadState = runtimeStore.GetThreadState(normalizedThreadId)
            If threadState Is Nothing OrElse threadState.TokenUsage Is Nothing Then
                UpdateTokenUsageWidget(normalizedThreadId, String.Empty, Nothing)
                Return
            End If

            Dim tokenTurnId = If(String.IsNullOrWhiteSpace(threadState.LatestTurnId),
                                 runtimeStore.GetLatestTurnId(normalizedThreadId),
                                 threadState.LatestTurnId)
            UpdateTokenUsageWidget(normalizedThreadId, tokenTurnId, threadState.TokenUsage)
        End Sub

        Private Shared Function ShouldSkipCompletedTurnSnapshotDuplicatedOverlayItem(item As TurnItemRuntimeState) As Boolean
            If item Is Nothing Then
                Return False
            End If

            Select Case If(item.ItemType, String.Empty).Trim().ToLowerInvariant()
                Case "usermessage", "agentmessage", "reasoning"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Sub AssignTurnItemOrderIndexes(items As IReadOnlyList(Of TurnItemRuntimeState))
            If items Is Nothing Then
                Return
            End If

            For i = 0 To items.Count - 1
                Dim item = items(i)
                If item Is Nothing Then
                    Continue For
                End If

                item.TurnItemOrderIndex = i
            Next
        End Sub

        Private Shared Function OrderCompletedTurnRuntimeOnlyReplayItems(items As IReadOnlyList(Of TurnItemRuntimeState)) As IReadOnlyList(Of TurnItemRuntimeState)
            If items Is Nothing OrElse items.Count <= 1 Then
                Return If(items, New List(Of TurnItemRuntimeState)())
            End If

            Dim indexedItems As New List(Of KeyValuePair(Of Integer, TurnItemRuntimeState))(items.Count)
            For i = 0 To items.Count - 1
                indexedItems.Add(New KeyValuePair(Of Integer, TurnItemRuntimeState)(i, items(i)))
            Next

            indexedItems.Sort(
                Function(left, right)
                    Dim streamSequenceCompare = CompareCompletedTurnReplayItemStreamSequence(left.Value, right.Value)
                    If streamSequenceCompare <> 0 Then
                        Return streamSequenceCompare
                    End If

                    Dim timestampCompare = CompareCompletedTurnReplayItemTimestamp(left.Value, right.Value)
                    If timestampCompare <> 0 Then
                        Return timestampCompare
                    End If

                    Dim typeCompare = CompareCompletedTurnReplayItemTypePriority(left.Value, right.Value)
                    If typeCompare <> 0 Then
                        Return typeCompare
                    End If

                    Return left.Key.CompareTo(right.Key)
                End Function)

            Dim results As New List(Of TurnItemRuntimeState)(indexedItems.Count)
            For Each pair In indexedItems
                results.Add(pair.Value)
            Next

            Return results
        End Function

        Private Shared Function CompareCompletedTurnReplayItemStreamSequence(left As TurnItemRuntimeState,
                                                                             right As TurnItemRuntimeState) As Integer
            Dim leftSequence = left?.TurnItemStreamSequence
            Dim rightSequence = right?.TurnItemStreamSequence

            If leftSequence.HasValue AndAlso rightSequence.HasValue Then
                Return leftSequence.Value.CompareTo(rightSequence.Value)
            End If

            If leftSequence.HasValue Then
                Return -1
            End If

            If rightSequence.HasValue Then
                Return 1
            End If

            Return 0
        End Function

        Private Shared Function CompareCompletedTurnReplayItemTimestamp(left As TurnItemRuntimeState,
                                                                        right As TurnItemRuntimeState) As Integer
            Dim leftTimestamp = GetCompletedTurnReplayItemSortTimestamp(left)
            Dim rightTimestamp = GetCompletedTurnReplayItemSortTimestamp(right)

            If leftTimestamp.HasValue AndAlso rightTimestamp.HasValue Then
                Return DateTimeOffset.Compare(leftTimestamp.Value, rightTimestamp.Value)
            End If

            If leftTimestamp.HasValue Then
                Return -1
            End If

            If rightTimestamp.HasValue Then
                Return 1
            End If

            Return 0
        End Function

        Private Shared Function GetCompletedTurnReplayItemSortTimestamp(item As TurnItemRuntimeState) As DateTimeOffset?
            If item Is Nothing Then
                Return Nothing
            End If

            If item.StartedAt.HasValue Then
                Return item.StartedAt.Value
            End If

            If item.CompletedAt.HasValue Then
                Return item.CompletedAt.Value
            End If

            Return Nothing
        End Function

        Private Shared Function CompareCompletedTurnReplayItemTypePriority(left As TurnItemRuntimeState,
                                                                           right As TurnItemRuntimeState) As Integer
            Return GetCompletedTurnReplayItemTypePriority(left).CompareTo(GetCompletedTurnReplayItemTypePriority(right))
        End Function

        Private Shared Function GetCompletedTurnReplayItemTypePriority(item As TurnItemRuntimeState) As Integer
            Select Case If(item?.ItemType, String.Empty).Trim().ToLowerInvariant()
                Case "commandexecution"
                    Return 0
                Case "filechange"
                    Return 1
                Case Else
                    Return 2
            End Select
        End Function

        Private Function ResolveOverlayTurnIdsForReplay(threadId As String,
                                                        runtimeStore As TurnFlowRuntimeStore) As HashSet(Of String)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            Dim results As New HashSet(Of String)(StringComparer.Ordinal)
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse runtimeStore Is Nothing Then
                Return results
            End If

            For Each turnId In _threadLiveSessionRegistry.GetOverlayTurnIds(normalizedThreadId)
                Dim normalizedTurnId = If(turnId, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                    Continue For
                End If

                results.Add(normalizedTurnId)
            Next

            Dim activeTurnId = runtimeStore.GetActiveTurnId(normalizedThreadId)
            If Not String.IsNullOrWhiteSpace(activeTurnId) Then
                results.Add(activeTurnId)
            End If

            Return results
        End Function

        Private Function EnumerateRuntimeItemsForThread(threadId As String) As IReadOnlyList(Of TurnItemRuntimeState)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return New List(Of TurnItemRuntimeState)()
            End If

            Return _sessionNotificationCoordinator.RuntimeStore.GetOrderedRuntimeItemsForThread(normalizedThreadId)
        End Function

        Private Function RemoveProjectionSnapshotItemRowsForCompletedOverlayTurns(threadId As String,
                                                                                  projection As ThreadTranscriptProjectionSnapshot,
                                                                                  runtimeStore As TurnFlowRuntimeStore) As HashSet(Of String)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            Dim completedTurnIds As New HashSet(Of String)(StringComparer.Ordinal)
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse projection Is Nothing OrElse runtimeStore Is Nothing Then
                Return completedTurnIds
            End If

            For Each turnId In ResolveOverlayTurnIdsForReplay(normalizedThreadId, runtimeStore)
                Dim normalizedTurnId = If(turnId, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                    Continue For
                End If

                Dim turnState = runtimeStore.GetTurnState(normalizedThreadId, normalizedTurnId)
                If turnState Is Nothing OrElse Not turnState.IsCompleted Then
                    Continue For
                End If

                completedTurnIds.Add(normalizedTurnId)
            Next

            If completedTurnIds.Count = 0 OrElse projection.DisplayEntries Is Nothing OrElse projection.DisplayEntries.Count = 0 Then
                Return completedTurnIds
            End If

            For i = projection.DisplayEntries.Count - 1 To 0 Step -1
                Dim descriptor = projection.DisplayEntries(i)
                If descriptor Is Nothing Then
                    Continue For
                End If

                Dim descriptorTurnId = If(descriptor.TurnId, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(descriptorTurnId) OrElse Not completedTurnIds.Contains(descriptorTurnId) Then
                    Continue For
                End If

                Dim runtimeKey = If(descriptor.RuntimeKey, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(runtimeKey) OrElse
                   Not runtimeKey.StartsWith("item:", StringComparison.Ordinal) Then
                    Continue For
                End If

                projection.DisplayEntries.RemoveAt(i)
            Next

            Return completedTurnIds
        End Function

        Private Sub RemoveProjectionOverlayReplaySupersededMarkers(threadId As String,
                                                                   projection As ThreadTranscriptProjectionSnapshot,
                                                                   runtimeStore As TurnFlowRuntimeStore)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse projection Is Nothing OrElse runtimeStore Is Nothing OrElse
               projection.DisplayEntries Is Nothing OrElse projection.DisplayEntries.Count = 0 Then
                Return
            End If

            Dim overlayTurnIds = ResolveOverlayTurnIdsForReplay(normalizedThreadId, runtimeStore)
            If overlayTurnIds.Count = 0 Then
                Return
            End If

            For i = projection.DisplayEntries.Count - 1 To 0 Step -1
                Dim descriptor = projection.DisplayEntries(i)
                If descriptor Is Nothing Then
                    Continue For
                End If

                Dim descriptorTurnId = If(descriptor.TurnId, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(descriptorTurnId) OrElse Not overlayTurnIds.Contains(descriptorTurnId) Then
                    Continue For
                End If

                Dim runtimeKey = If(descriptor.RuntimeKey, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(runtimeKey) Then
                    Continue For
                End If

                Dim isTurnStartMarker = StringComparer.Ordinal.Equals(runtimeKey, $"turn:lifecycle:start:{descriptorTurnId}")
                Dim isTurnPlanMarker = runtimeKey.StartsWith("turn:meta:plan:", StringComparison.Ordinal) AndAlso
                                       runtimeKey.EndsWith(":" & descriptorTurnId, StringComparison.Ordinal)
                Dim isTurnDiffMarker = runtimeKey.StartsWith("turn:meta:diff:", StringComparison.Ordinal) AndAlso
                                       runtimeKey.EndsWith(":" & descriptorTurnId, StringComparison.Ordinal)

                If Not isTurnStartMarker AndAlso Not isTurnPlanMarker AndAlso Not isTurnDiffMarker Then
                    Continue For
                End If

                projection.DisplayEntries.RemoveAt(i)
            Next
        End Sub

        Private Shared Function ThreadObjectHasTurns(threadObject As JsonObject) As Boolean
            Dim turns = GetPropertyArray(threadObject, "turns")
            Return turns IsNot Nothing AndAlso turns.Count > 0
        End Function

        Private Function BuildThreadTranscriptSnapshot(threadObject As JsonObject) As ThreadTranscriptSnapshot
            Dim assistantPhaseHintLookup = CreateSnapshotAssistantPhaseHintLookupCopy()
            Return ThreadTranscriptSnapshotBuilder.BuildFromThread(threadObject, assistantPhaseHintLookup)
        End Function

        Private Function CreateSnapshotAssistantPhaseHintLookupCopy() As Dictionary(Of String, String)
            SyncLock _snapshotAssistantPhaseHintsLock
                Return New Dictionary(Of String, String)(_snapshotAssistantPhaseHintsByItemKey, StringComparer.Ordinal)
            End SyncLock
        End Function

        Private Sub MergeSnapshotAssistantPhaseHints(hints As IReadOnlyDictionary(Of String, String))
            If hints Is Nothing OrElse hints.Count = 0 Then
                Return
            End If

            SyncLock _snapshotAssistantPhaseHintsLock
                For Each pair In hints
                    Dim key = If(pair.Key, String.Empty).Trim()
                    Dim phase = If(pair.Value, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(key) OrElse String.IsNullOrWhiteSpace(phase) Then
                        Continue For
                    End If

                    _snapshotAssistantPhaseHintsByItemKey(key) = phase
                Next
            End SyncLock
        End Sub

        Private Sub CacheAssistantPhaseHintFromRuntimeItem(itemState As TurnItemRuntimeState)
            If itemState Is Nothing Then
                Return
            End If

            If Not StringComparer.OrdinalIgnoreCase.Equals(itemState.ItemType, "agentMessage") Then
                Return
            End If

            Dim phase = If(itemState.AgentMessagePhase, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(phase) Then
                Return
            End If

            Dim threadId = If(itemState.ThreadId, String.Empty).Trim()
            Dim turnId = If(itemState.TurnId, String.Empty).Trim()
            Dim itemId = If(itemState.ItemId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(threadId) OrElse
               String.IsNullOrWhiteSpace(turnId) OrElse
               String.IsNullOrWhiteSpace(itemId) Then
                Return
            End If

            Dim key = $"{threadId}:{turnId}:{itemId}"
            SyncLock _snapshotAssistantPhaseHintsLock
                _snapshotAssistantPhaseHintsByItemKey(key) = phase
            End SyncLock
        End Sub

        Private Sub ApplyCurrentThreadFromThreadObject(threadObject As JsonObject,
                                                       Optional clearPendingNewThreadSelection As Boolean = True)
            If clearPendingNewThreadSelection Then
                SetPendingNewThreadFirstPromptSelectionActive(False)
            End If

            Dim threadId = GetPropertyString(threadObject, "id")
            If Not String.IsNullOrWhiteSpace(threadId) Then
                SetVisibleThreadId(threadId)
            End If

            Dim loadedThreadCwd = ExtractThreadWorkingDirectoryFromThreadObject(threadObject)
            If Not String.IsNullOrWhiteSpace(loadedThreadCwd) Then
                _currentThreadCwd = loadedThreadCwd
                _newThreadTargetOverrideCwd = String.Empty
            ElseIf String.IsNullOrWhiteSpace(_newThreadTargetOverrideCwd) Then
                _currentThreadCwd = String.Empty
            End If

            ClearVisibleTurnId()
            SyncCurrentTurnFromRuntimeStore(keepExistingWhenRuntimeIsIdle:=False)
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
