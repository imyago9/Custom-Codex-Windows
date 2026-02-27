Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Threading
Imports CodexNativeAgent.AppServer
Imports CodexNativeAgent.Services
Imports CodexNativeAgent.Ui.Coordinators
Imports CodexNativeAgent.Ui.ViewModels

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private Const TranscriptChunkSessionDebugInstrumentationEnabled As Boolean = False
        Private Const TranscriptDocumentCacheMaxInactiveEntries As Integer = 3
        Private Const RateLimitAutoRefreshMinIntervalSeconds As Integer = 45
        Private ReadOnly _threadTranscriptChunkSessionCoordinator As New ThreadTranscriptChunkSessionCoordinator()
        Private ReadOnly _inactiveTranscriptDocumentsByThreadId As New Dictionary(Of String, CachedTranscriptDocumentState)(StringComparer.Ordinal)
        Private ReadOnly _rateLimitStatesByLimitId As New Dictionary(Of String, RateLimitLimitState)(StringComparer.OrdinalIgnoreCase)
        Private _activeTranscriptDocumentThreadId As String = String.Empty
        Private _activeTranscriptDocumentContentFingerprint As String = String.Empty
        Private _activeTranscriptDocumentUpdatedAtUnix As Long = Long.MinValue
        Private _rateLimitAutoRefreshInProgress As Boolean
        Private _lastRateLimitAutoRefreshAttemptUtc As DateTimeOffset = DateTimeOffset.MinValue

        Private NotInheritable Class CachedTranscriptDocumentState
            Public Property ThreadId As String = String.Empty
            Public Property State As TranscriptPanelViewModel.TranscriptThreadDocumentState
            Public Property LastUsedUtc As DateTimeOffset = DateTimeOffset.UtcNow
            Public Property ContentFingerprint As String = String.Empty
            Public Property ThreadUpdatedAtUnix As Long = Long.MinValue
        End Class

        Private NotInheritable Class RateLimitBucketState
            Public Property UsedPercent As Double?
            Public Property WindowDurationMins As Integer?
            Public Property ResetsAtUnix As Long?
        End Class

        Private NotInheritable Class RateLimitLimitState
            Public Property LimitId As String = String.Empty
            Public Property LimitName As String = String.Empty
            Public Property Primary As RateLimitBucketState
            Public Property Secondary As RateLimitBucketState
        End Class

        ' Visible selection helpers (Phase 7 scaffold):
        ' `_currentThreadId` / `_currentTurnId` represent the currently selected thread/turn in the UI.
        ' They are not the global source of truth for active runtime state across all threads.
        ' Runtime truth lives in TurnFlowRuntimeStore + ThreadLiveSessionRegistry.
        ' Prefer these helpers over direct field access so visible-vs-global semantics stay explicit.
        Private Function GetVisibleThreadId() As String
            Return If(_currentThreadId, String.Empty).Trim()
        End Function

        Private Sub SetVisibleThreadId(value As String)
            Dim normalizedValue = If(value, String.Empty).Trim()
            Dim previousThreadId = GetVisibleThreadId()

            If StringComparer.Ordinal.Equals(previousThreadId, normalizedValue) Then
                If Not String.IsNullOrWhiteSpace(normalizedValue) Then
                    _threadTranscriptChunkSessionCoordinator.ActivateVisibleThread(normalizedValue, "visible_thread_reaffirmed")
                    TraceTranscriptChunkSession("visible_thread_reaffirmed")
                End If

                Return
            End If

            SyncTurnComposerStateForCurrentSelection()
            _currentThreadId = normalizedValue

            If String.IsNullOrWhiteSpace(normalizedValue) Then
                Dim removedSession = _threadTranscriptChunkSessionCoordinator.ResetActiveSession("visible_thread_cleared")
                TraceTranscriptChunkSession("visible_thread_cleared", $"previous={previousThreadId}", removedSession)
                RestoreTurnComposerStateForCurrentSelection("visible_thread_cleared")
                Return
            End If

            _threadTranscriptChunkSessionCoordinator.ActivateVisibleThread(normalizedValue, "visible_thread_changed")
            TraceTranscriptChunkSession("visible_thread_changed", $"previous={previousThreadId}")
            RestoreTurnComposerStateForCurrentSelection("visible_thread_changed")
        End Sub

        Private Sub ClearVisibleThreadId()
            SetVisibleThreadId(String.Empty)
        End Sub

        Private Function GetVisibleTurnId() As String
            Return If(_currentTurnId, String.Empty).Trim()
        End Function

        Private Sub SetVisibleTurnId(value As String)
            _currentTurnId = If(value, String.Empty).Trim()
        End Sub

        Private Sub ClearVisibleTurnId()
            _currentTurnId = String.Empty
        End Sub

        Private Sub ClearVisibleSelection()
            ClearVisibleThreadId()
            ClearVisibleTurnId()
        End Sub

        Private Function EnsureTranscriptDocumentActivatedForThread(threadId As String,
                                                                   Optional reason As String = Nothing) As Boolean
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return ActivateFreshTranscriptDocument(reason)
            End If

            If _viewModel Is Nothing OrElse _viewModel.TranscriptPanel Is Nothing Then
                Return False
            End If

            If StringComparer.Ordinal.Equals(_activeTranscriptDocumentThreadId, normalizedThreadId) Then
                TouchCachedTranscriptDocument(normalizedThreadId)
                Return False
            End If

            Dim nextState As TranscriptPanelViewModel.TranscriptThreadDocumentState = Nothing
            Dim nextContentFingerprint As String = String.Empty
            Dim nextThreadUpdatedAtUnix As Long = Long.MinValue
            Dim reusedCachedState = TryTakeCachedTranscriptDocumentState(normalizedThreadId,
                                                                         nextState,
                                                                         nextContentFingerprint,
                                                                         nextThreadUpdatedAtUnix)
            If nextState Is Nothing Then
                nextState = TranscriptPanelViewModel.CreateEmptyThreadDocumentState()
                nextContentFingerprint = String.Empty
                nextThreadUpdatedAtUnix = Long.MinValue
            End If

            Dim previousThreadId = _activeTranscriptDocumentThreadId
            Dim swapPerf = Stopwatch.StartNew()
            Dim previousState = _viewModel.TranscriptPanel.SwapThreadDocumentState(nextState)
            Dim swapMs = swapPerf.ElapsedMilliseconds

            If Not String.IsNullOrWhiteSpace(previousThreadId) AndAlso previousState IsNot Nothing Then
                StoreCachedTranscriptDocumentState(previousThreadId,
                                                  previousState,
                                                  _activeTranscriptDocumentContentFingerprint,
                                                  _activeTranscriptDocumentUpdatedAtUnix)
            End If

            _activeTranscriptDocumentThreadId = normalizedThreadId
            _activeTranscriptDocumentContentFingerprint = If(nextContentFingerprint, String.Empty).Trim()
            _activeTranscriptDocumentUpdatedAtUnix = nextThreadUpdatedAtUnix
            TrimInactiveTranscriptDocuments()
            EnsureTranscriptTabSurfaceActivatedForThread(normalizedThreadId)

            AppendProtocol("debug",
                           $"transcript_doc_swap thread={normalizedThreadId} previous={previousThreadId} reused={reusedCachedState} swapMs={swapMs} inactiveCached={_inactiveTranscriptDocumentsByThreadId.Count} reason={If(reason, String.Empty)}")
            Return True
        End Function

        Private Function ActivateFreshTranscriptDocument(Optional reason As String = Nothing,
                                                        Optional activateBlankSurface As Boolean = True) As Boolean
            If _viewModel Is Nothing OrElse _viewModel.TranscriptPanel Is Nothing Then
                Return False
            End If

            Dim previousThreadId = _activeTranscriptDocumentThreadId
            Dim hasPreviousThread = Not String.IsNullOrWhiteSpace(previousThreadId)
            Dim nextState = TranscriptPanelViewModel.CreateEmptyThreadDocumentState()

            If Not hasPreviousThread AndAlso _viewModel.TranscriptPanel.Items.Count = 0 Then
                Return False
            End If

            Dim swapPerf = Stopwatch.StartNew()
            Dim previousState = _viewModel.TranscriptPanel.SwapThreadDocumentState(nextState)
            Dim swapMs = swapPerf.ElapsedMilliseconds

            If hasPreviousThread AndAlso previousState IsNot Nothing Then
                StoreCachedTranscriptDocumentState(previousThreadId,
                                                  previousState,
                                                  _activeTranscriptDocumentContentFingerprint,
                                                  _activeTranscriptDocumentUpdatedAtUnix)
            End If

            _activeTranscriptDocumentThreadId = String.Empty
            _activeTranscriptDocumentContentFingerprint = String.Empty
            _activeTranscriptDocumentUpdatedAtUnix = Long.MinValue
            TrimInactiveTranscriptDocuments()
            If activateBlankSurface Then
                EnsureTranscriptTabSurfaceActivatedForThread(String.Empty)
            End If

            AppendProtocol("debug",
                           $"transcript_doc_swap_blank previous={previousThreadId} swapMs={swapMs} inactiveCached={_inactiveTranscriptDocumentsByThreadId.Count} activateBlankSurface={activateBlankSurface} reason={If(reason, String.Empty)}")
            Return True
        End Function

        Private Sub StoreCachedTranscriptDocumentState(threadId As String,
                                                       state As TranscriptPanelViewModel.TranscriptThreadDocumentState,
                                                       Optional contentFingerprint As String = Nothing,
                                                       Optional threadUpdatedAtUnix As Long = Long.MinValue)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse state Is Nothing Then
                Return
            End If

            _inactiveTranscriptDocumentsByThreadId(normalizedThreadId) = New CachedTranscriptDocumentState() With {
                .ThreadId = normalizedThreadId,
                .State = state,
                .LastUsedUtc = DateTimeOffset.UtcNow,
                .ContentFingerprint = If(contentFingerprint, String.Empty).Trim(),
                .ThreadUpdatedAtUnix = threadUpdatedAtUnix
            }
        End Sub

        Private Function TryTakeCachedTranscriptDocumentState(threadId As String,
                                                              ByRef state As TranscriptPanelViewModel.TranscriptThreadDocumentState,
                                                              ByRef contentFingerprint As String,
                                                              ByRef threadUpdatedAtUnix As Long) As Boolean
            state = Nothing
            contentFingerprint = String.Empty
            threadUpdatedAtUnix = Long.MinValue
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return False
            End If

            Dim entry As CachedTranscriptDocumentState = Nothing
            If Not _inactiveTranscriptDocumentsByThreadId.TryGetValue(normalizedThreadId, entry) OrElse entry Is Nothing Then
                Return False
            End If

            _inactiveTranscriptDocumentsByThreadId.Remove(normalizedThreadId)
            entry.LastUsedUtc = DateTimeOffset.UtcNow
            state = entry.State
            contentFingerprint = If(entry.ContentFingerprint, String.Empty).Trim()
            threadUpdatedAtUnix = entry.ThreadUpdatedAtUnix
            Return state IsNot Nothing
        End Function

        Private Sub TouchCachedTranscriptDocument(threadId As String)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            Dim entry As CachedTranscriptDocumentState = Nothing
            If _inactiveTranscriptDocumentsByThreadId.TryGetValue(normalizedThreadId, entry) AndAlso entry IsNot Nothing Then
                entry.LastUsedUtc = DateTimeOffset.UtcNow
            End If
        End Sub

        Private Sub TrimInactiveTranscriptDocuments()
            If _inactiveTranscriptDocumentsByThreadId.Count <= TranscriptDocumentCacheMaxInactiveEntries Then
                Return
            End If

            Do While _inactiveTranscriptDocumentsByThreadId.Count > TranscriptDocumentCacheMaxInactiveEntries
                Dim oldestKey As String = String.Empty
                Dim oldestUtc = DateTimeOffset.MaxValue

                For Each kvp In _inactiveTranscriptDocumentsByThreadId
                    Dim candidate = kvp.Value
                    If candidate Is Nothing Then
                        oldestKey = kvp.Key
                        Exit For
                    End If

                    If candidate.LastUsedUtc < oldestUtc Then
                        oldestUtc = candidate.LastUsedUtc
                        oldestKey = kvp.Key
                    End If
                Next

                If String.IsNullOrWhiteSpace(oldestKey) Then
                    Exit Do
                End If

                _inactiveTranscriptDocumentsByThreadId.Remove(oldestKey)
                RemoveRetainedTranscriptTabSurface(oldestKey)
                AppendProtocol("debug",
                               $"transcript_doc_retire thread={oldestKey} inactiveCached={_inactiveTranscriptDocumentsByThreadId.Count}")
            Loop
        End Sub

        Private Sub ClearCachedTranscriptDocuments()
            _inactiveTranscriptDocumentsByThreadId.Clear()
            _activeTranscriptDocumentThreadId = String.Empty
            _activeTranscriptDocumentContentFingerprint = String.Empty
            _activeTranscriptDocumentUpdatedAtUnix = Long.MinValue
            ClearRetainedTranscriptTabSurfaces()
        End Sub

        Private Function TryGetActiveTranscriptDocumentContentFingerprint(threadId As String,
                                                                          ByRef contentFingerprint As String,
                                                                          ByRef threadUpdatedAtUnix As Long) As Boolean
            contentFingerprint = String.Empty
            threadUpdatedAtUnix = Long.MinValue

            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return False
            End If

            If Not StringComparer.Ordinal.Equals(_activeTranscriptDocumentThreadId, normalizedThreadId) Then
                Return False
            End If

            contentFingerprint = If(_activeTranscriptDocumentContentFingerprint, String.Empty).Trim()
            threadUpdatedAtUnix = _activeTranscriptDocumentUpdatedAtUnix
            Return Not String.IsNullOrWhiteSpace(contentFingerprint)
        End Function

        Private Sub SetActiveTranscriptDocumentContentFingerprint(threadId As String,
                                                                 contentFingerprint As String,
                                                                 threadUpdatedAtUnix As Long)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            If Not StringComparer.Ordinal.Equals(_activeTranscriptDocumentThreadId, normalizedThreadId) Then
                Return
            End If

            _activeTranscriptDocumentContentFingerprint = If(contentFingerprint, String.Empty).Trim()
            _activeTranscriptDocumentUpdatedAtUnix = threadUpdatedAtUnix
        End Sub

        Private Function TryShowCachedTranscriptTabSelectionPreview(threadId As String,
                                                                    Optional reason As String = Nothing) As Boolean
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return False
            End If

            If Not HasRetainedTranscriptTabSurface(normalizedThreadId) Then
                Return False
            End If

            If Not StringComparer.Ordinal.Equals(_activeTranscriptDocumentThreadId, normalizedThreadId) AndAlso
               Not _inactiveTranscriptDocumentsByThreadId.ContainsKey(normalizedThreadId) Then
                Return False
            End If

            EnsureTranscriptDocumentActivatedForThread(normalizedThreadId, If(reason, "selection_cached_preview"))
            SetVisibleThreadId(normalizedThreadId)
            ClearVisibleTurnId()
            UpdateThreadTurnLabels()
            UpdateWorkspaceEmptyStateVisibility()
            RefreshControlStates()
            Return True
        End Function

        Private Sub TraceTranscriptChunkSession(eventName As String,
                                                Optional details As String = Nothing,
                                                Optional sessionOverride As ThreadTranscriptChunkSession = Nothing)
            If Not TranscriptChunkSessionDebugInstrumentationEnabled Then
                Return
            End If

            If _viewModel Is Nothing OrElse _viewModel.TranscriptPanel Is Nothing Then
                Return
            End If

            Dim session = If(sessionOverride, _threadTranscriptChunkSessionCoordinator.ActiveSession)
            Dim sb As New StringBuilder()
            sb.Append("transcript_chunk_session event=").Append(If(eventName, String.Empty))

            If Not String.IsNullOrWhiteSpace(details) Then
                sb.Append(" details=""").Append(details.Replace("""", "'")).Append(""""c)
            End If

            sb.Append(" visibleThread=").Append(GetVisibleThreadId())

            If session Is Nothing Then
                sb.Append(" active=false")
                AppendProtocol("debug", sb.ToString())
                Return
            End If

            sb.Append(" active=true")
            sb.Append(" thread=").Append(If(session.ThreadId, String.Empty))
            sb.Append(" gen=").Append(session.GenerationId)
            sb.Append(" loadingOlder=").Append(session.IsLoadingOlderChunk)
            sb.Append(" hasMoreOlder=").Append(session.HasMoreOlderChunks)
            sb.Append(" pendingPrepend=").Append(session.PendingPrependRequest)
            sb.Append(" loadsReq=").Append(session.OlderChunkLoadsRequested)
            sb.Append(" loadsDone=").Append(session.OlderChunkLoadsCompleted)
            sb.Append(" loadsCanceled=").Append(session.OlderChunkLoadsCanceled)

            If session.LoadedRangeStart.HasValue Then
                sb.Append(" rangeStart=").Append(session.LoadedRangeStart.Value)
            End If

            If session.LoadedRangeEnd.HasValue Then
                sb.Append(" rangeEnd=").Append(session.LoadedRangeEnd.Value)
            End If

            AppendProtocol("debug", sb.ToString())
        End Sub

        Private Function GetActiveTurnIdForThread(threadId As String,
                                                  Optional fallbackTurnId As String = Nothing) As String
            If _sessionNotificationCoordinator Is Nothing Then
                Return String.Empty
            End If

            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return String.Empty
            End If

            Dim normalizedFallbackTurnId = If(fallbackTurnId, String.Empty).Trim()
            Return If(_sessionNotificationCoordinator.RuntimeStore.GetActiveTurnId(normalizedThreadId, normalizedFallbackTurnId),
                      String.Empty).Trim()
        End Function

        Private Function HasActiveRuntimeTurnForThread(threadId As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(GetActiveTurnIdForThread(threadId))
        End Function

        Private Sub UpdateRuntimeFieldState()
            Dim connected = IsClientRunning()
            Dim editable = Not connected

            _viewModel.SettingsPanel.AreConnectionFieldsEditable = editable
        End Sub

        Private Sub UpdateConnectionStateTextFromSession(Optional syncFirst As Boolean = True)
            If syncFirst Then
                SyncSessionStateViewModel()
            End If

            _viewModel.ConnectionStateText = _viewModel.SessionState.ConnectionStateLabel
        End Sub

        Private Sub SyncSessionStateViewModel()
            SyncCurrentTurnFromRuntimeStore(keepExistingWhenRuntimeIsIdle:=True)
            Dim session = _viewModel.SessionState
            session.IsConnected = IsClientRunning()
            session.IsAuthenticated = _isAuthenticated
            session.ConnectionExpected = _connectionExpected
            session.IsReconnectInProgress = _reconnectInProgress
            session.ReconnectAttempt = _reconnectAttempt
            session.NextReconnectAttemptUtc = _nextReconnectAttemptUtc
            session.LastActivityUtc = _lastActivityUtc
            session.CurrentLoginId = _currentLoginId
            session.CurrentThreadId = GetVisibleThreadId()
            session.CurrentTurnId = GetVisibleTurnId()
            session.ProcessId = If(_client Is Nothing, 0, _client.ProcessId)
        End Sub

        Private Sub SyncCurrentTurnFromRuntimeStore(Optional keepExistingWhenRuntimeIsIdle As Boolean = False)
            If _sessionNotificationCoordinator Is Nothing Then
                Return
            End If

            Dim normalizedThreadId = GetVisibleThreadId()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                If Not keepExistingWhenRuntimeIsIdle Then
                    ClearVisibleTurnId()
                End If
                Return
            End If

            Dim runtimeStore = _sessionNotificationCoordinator.RuntimeStore
            Dim activeTurnId = GetActiveTurnIdForThread(normalizedThreadId, GetVisibleTurnId())
            If Not String.IsNullOrWhiteSpace(activeTurnId) Then
                SetVisibleTurnId(activeTurnId)
                Return
            End If

            If keepExistingWhenRuntimeIsIdle Then
                Dim latestTurnId = runtimeStore.GetLatestTurnId(normalizedThreadId)
                If String.IsNullOrWhiteSpace(latestTurnId) Then
                    Return
                End If
            End If

            ClearVisibleTurnId()
        End Sub

        Private Function HasActiveRuntimeTurnForCurrentThread() As Boolean
            Return HasActiveRuntimeTurnForThread(GetVisibleThreadId())
        End Function

        Private Sub OnTranscriptLeadingEntriesTrimmed(sender As Object,
                                                      e As TranscriptPanelViewModel.TranscriptLeadingEntriesTrimmedEventArgs)
            If e Is Nothing OrElse e.RemovedCount <= 0 Then
                Return
            End If

            Dim activeSession = _threadTranscriptChunkSessionCoordinator.ActiveSession
            If activeSession Is Nothing Then
                Return
            End If

            Dim visibleThreadId = GetVisibleThreadId()
            Dim sessionThreadId = If(activeSession.ThreadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(visibleThreadId) OrElse
               Not StringComparer.Ordinal.Equals(sessionThreadId, visibleThreadId) Then
                Return
            End If

            If activeSession.LoadedRangeStart.HasValue Then
                activeSession.LoadedRangeStart = Math.Max(0, activeSession.LoadedRangeStart.Value + e.RemovedCount)
            End If

            If activeSession.LoadedRangeEnd.HasValue AndAlso
               activeSession.LoadedRangeStart.HasValue AndAlso
               activeSession.LoadedRangeEnd.Value < activeSession.LoadedRangeStart.Value Then
                activeSession.LoadedRangeEnd = activeSession.LoadedRangeStart.Value
            End If

            activeSession.LastUpdatedUtc = DateTimeOffset.UtcNow
            activeSession.LastLifecycleReason = "leading_trim_adjust_range"

            TraceTranscriptChunkSession("leading_trim_adjust_range",
                                        $"thread={sessionThreadId}; removed={e.RemovedCount}; rangeStart={If(activeSession.LoadedRangeStart.HasValue, activeSession.LoadedRangeStart.Value.ToString(CultureInfo.InvariantCulture), "none")}; rangeEnd={If(activeSession.LoadedRangeEnd.HasValue, activeSession.LoadedRangeEnd.Value.ToString(CultureInfo.InvariantCulture), "none")}")
        End Sub

        Private Function RuntimeHasTurnHistoryForCurrentThread() As Boolean
            If _sessionNotificationCoordinator Is Nothing Then
                Return False
            End If

            Dim normalizedThreadId = GetVisibleThreadId()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return False
            End If

            Dim latestTurnId = _sessionNotificationCoordinator.RuntimeStore.GetLatestTurnId(normalizedThreadId)
            Return Not String.IsNullOrWhiteSpace(latestTurnId)
        End Function

        Private Sub SetSessionAuthenticated(value As Boolean)
            _isAuthenticated = value
            _viewModel.SessionState.IsAuthenticated = value
        End Sub

        Private Sub SetSessionConnectionExpected(value As Boolean)
            _connectionExpected = value
            _viewModel.SessionState.ConnectionExpected = value
        End Sub

        Private Sub SetSessionReconnectInProgress(value As Boolean)
            _reconnectInProgress = value
            _viewModel.SessionState.IsReconnectInProgress = value
        End Sub

        Private Sub SetSessionReconnectAttempt(value As Integer)
            _reconnectAttempt = value
            _viewModel.SessionState.ReconnectAttempt = value
        End Sub

        Private Sub SetSessionNextReconnectAttempt(value As DateTimeOffset?)
            _nextReconnectAttemptUtc = value
            _viewModel.SessionState.NextReconnectAttemptUtc = value
        End Sub

        Private Sub SetSessionLastActivity(value As DateTimeOffset)
            _lastActivityUtc = value
            _viewModel.SessionState.LastActivityUtc = value
        End Sub

        Private Sub SetSessionCurrentLoginId(value As String)
            _currentLoginId = If(value, String.Empty)
            _viewModel.SessionState.CurrentLoginId = _currentLoginId
        End Sub

        Private Sub InitializeSessionCoordinatorBindings()
            If _sessionCoordinator Is Nothing Then
                Return
            End If

            AddHandler _sessionCoordinator.ReconnectAttemptStarted, AddressOf OnSessionCoordinatorReconnectAttemptStarted
            AddHandler _sessionCoordinator.ReconnectAttemptFailed, AddressOf OnSessionCoordinatorReconnectAttemptFailed
            AddHandler _sessionCoordinator.ReconnectSucceeded, AddressOf OnSessionCoordinatorReconnectSucceeded
            AddHandler _sessionCoordinator.ReconnectRetryScheduled, AddressOf OnSessionCoordinatorReconnectRetryScheduled
            AddHandler _sessionCoordinator.ReconnectTerminalFailure, AddressOf OnSessionCoordinatorReconnectTerminalFailure
            AddHandler _sessionCoordinator.ReconnectCanceled, AddressOf OnSessionCoordinatorReconnectCanceled
            AddHandler _sessionCoordinator.ReconnectFinalizing, AddressOf OnSessionCoordinatorReconnectFinalizing
        End Sub

        Private Sub OnSessionCoordinatorReconnectAttemptStarted(sender As Object,
                                                               e As SessionCoordinator.ReconnectAttemptStartedEventArgs)
            If e Is Nothing Then
                Return
            End If

            RunOnUi(
                Sub()
                    ApplySessionReconnectAttemptStarted(e.CurrentAttempt)
                    NotifyReconnectAttemptStarted(e.Reason, e.CurrentAttempt, e.TotalAttempts)
                End Sub)
        End Sub

        Private Sub OnSessionCoordinatorReconnectAttemptFailed(sender As Object,
                                                              e As SessionCoordinator.ReconnectAttemptFailedEventArgs)
            If e Is Nothing Then
                Return
            End If

            RunOnUi(
                Sub() NotifyReconnectAttemptFailed(e.CurrentAttempt, e.Error))
        End Sub

        Private Sub OnSessionCoordinatorReconnectSucceeded(sender As Object, e As EventArgs)
            RunOnUi(
                Sub() NotifyReconnectSucceeded())
        End Sub

        Private Sub OnSessionCoordinatorReconnectRetryScheduled(sender As Object,
                                                                e As SessionCoordinator.ReconnectRetryScheduledEventArgs)
            If e Is Nothing Then
                Return
            End If

            RunOnUi(
                Sub()
                    ApplySessionReconnectRetryScheduled(e.NextAttemptUtc)
                    UpdateReconnectCountdownUi()
                End Sub)
        End Sub

        Private Sub OnSessionCoordinatorReconnectTerminalFailure(sender As Object, e As EventArgs)
            RunOnUi(
                Sub() HandleReconnectTerminalFailureUi())
        End Sub

        Private Sub OnSessionCoordinatorReconnectCanceled(sender As Object, e As EventArgs)
            RunOnUi(
                Sub() HandleReconnectCanceledUi())
        End Sub

        Private Sub OnSessionCoordinatorReconnectFinalizing(sender As Object,
                                                            e As SessionCoordinator.ReconnectFinalizeEventArgs)
            If e Is Nothing Then
                Return
            End If

            RunOnUi(
                Sub() FinalizeReconnectLoopUi(e.ReconnectCancellationSource))
        End Sub

        Private Sub ApplySessionConnectionEstablished(nowUtc As DateTimeOffset)
            _connectionExpected = True
            _lastActivityUtc = nowUtc
            _nextReconnectAttemptUtc = Nothing
            _viewModel.SessionState.ApplyConnectionEstablished(nowUtc)
            UpdateConnectionStateTextFromSession()
        End Sub

        Private Sub ApplySessionReconnectLoopStartedState()
            _connectionExpected = True
            _reconnectInProgress = True
            _reconnectAttempt = 0
            _nextReconnectAttemptUtc = Nothing
            _viewModel.SessionState.ApplyReconnectLoopStarted()
        End Sub

        Private Sub ApplySessionReconnectAttemptStarted(attempt As Integer)
            _reconnectAttempt = Math.Max(0, attempt)
            _nextReconnectAttemptUtc = Nothing
            _viewModel.SessionState.ApplyReconnectAttemptStarted(attempt)
        End Sub

        Private Sub ApplySessionReconnectRetryScheduled(nextAttemptUtc As DateTimeOffset)
            _nextReconnectAttemptUtc = nextAttemptUtc
            _viewModel.SessionState.ApplyReconnectRetryScheduled(nextAttemptUtc)
        End Sub

        Private Sub ApplySessionReconnectTerminalFailureState()
            _connectionExpected = False
            _nextReconnectAttemptUtc = Nothing
            _viewModel.SessionState.ApplyReconnectTerminalFailureState()
        End Sub

        Private Sub ApplySessionReconnectCanceledState()
            _nextReconnectAttemptUtc = Nothing
            _viewModel.SessionState.ApplyReconnectCanceledState()
        End Sub

        Private Sub ApplySessionReconnectIdleState()
            _reconnectInProgress = False
            _reconnectAttempt = 0
            _nextReconnectAttemptUtc = Nothing
            _viewModel.SessionState.ApplyReconnectIdleState()
        End Sub

        Private Sub ClearAuthRequiredNoticeState()
            _authRequiredNoticeShown = False
        End Sub

        Private Sub AppendAndShowSystemMessage(message As String,
                                               Optional isError As Boolean = False,
                                               Optional displayToast As Boolean = False)
            If String.IsNullOrWhiteSpace(message) Then
                Return
            End If

            AppendSystemMessage(message)
            ShowStatus(message, isError:=isError, displayToast:=displayToast)
        End Sub

        Private Sub AppendAndShowSystemMessage(logMessage As String,
                                               statusMessage As String,
                                               Optional isError As Boolean = False,
                                               Optional displayToast As Boolean = False)
            If Not String.IsNullOrWhiteSpace(logMessage) Then
                AppendSystemMessage(logMessage)
            End If

            If Not String.IsNullOrWhiteSpace(statusMessage) Then
                ShowStatus(statusMessage, isError:=isError, displayToast:=displayToast)
            End If
        End Sub

        Private Sub ApplyAccountAuthenticationPresence(hasAccount As Boolean)
            SetSessionAuthenticated(hasAccount)
            If hasAccount Then
                ClearAuthRequiredNoticeState()
            End If
        End Sub

        Private Function BuildAccountStateText(result As AccountReadResult) As String
            If result Is Nothing Then
                Return "Account: unknown"
            End If

            If result.Account IsNot Nothing Then
                Return FormatAccountLabel(result.Account, result.RequiresOpenAiAuth)
            End If

            Return If(result.RequiresOpenAiAuth,
                      "Account: signed out (OpenAI auth required)",
                      "Account: unknown")
        End Function

        Private Function ApplyAccountReadResultUi(result As AccountReadResult) As Boolean
            _viewModel.SettingsPanel.AccountStateText = BuildAccountStateText(result)
            Dim hasAccount = result IsNot Nothing AndAlso result.Account IsNot Nothing
            ApplyAccountAuthenticationPresence(hasAccount)
            RefreshControlStates()
            Return hasAccount
        End Function

        Private Function ShouldInitializeWorkspaceAfterAuthentication(wasAuthenticated As Boolean) As Boolean
            Return (Not wasAuthenticated) OrElse
                   (Not _modelsLoadedAtLeastOnce) OrElse
                   (Not _threadsLoadedAtLeastOnce)
        End Function

        Private Function CanBeginWorkspaceBootstrapAfterAuthentication() As Boolean
            SyncSessionStateViewModel()
            If Not _viewModel.SessionState.IsConnectedAndAuthenticated Then
                Return False
            End If

            If _workspaceBootstrapInProgress Then
                Return False
            End If

            _workspaceBootstrapInProgress = True
            Return True
        End Function

        Private Sub EndWorkspaceBootstrapAfterAuthentication()
            _workspaceBootstrapInProgress = False
        End Sub

        Private Sub ResetWorkspaceTransientStateCore(clearModelPicker As Boolean)
            ActivateFreshTranscriptDocument("workspace_reset")
            ClearCachedTranscriptDocuments()
            ClearVisibleSelection()
            SetPendingNewThreadFirstPromptSelectionActive(False, clearThreadSelection:=False)
            _notificationRuntimeThreadId = String.Empty
            _notificationRuntimeTurnId = String.Empty
            _currentThreadCwd = String.Empty
            _newThreadTargetOverrideCwd = String.Empty
            _threadsLoading = False
            _viewModel.ThreadsPanel.IsLoading = False
            _viewModel.ThreadsPanel.RefreshErrorText = String.Empty
            ResetThreadSelectionLoadUiState()
            _modelsLoadedAtLeastOnce = False
            _threadsLoadedAtLeastOnce = False
            _workspaceBootstrapInProgress = False
            _threadEntries.Clear()
            _expandedThreadProjectGroups.Clear()
            ClearTurnComposerThreadStates()
            _threadLiveSessionRegistry.Clear()
            _sessionNotificationCoordinator.ResetStreamingAgentItems()
            _pendingLocalUserEchoes.Clear()
            _turnWorkflowCoordinator.ResetApprovalState()
            _viewModel.ThreadsPanel.Items.Clear()
            _viewModel.ThreadsPanel.SelectedListItem = Nothing
            ClearTurnComposerRateLimitStateUi()
            If clearModelPicker Then
                WorkspacePaneHost.CmbModel.Items.Clear()
            End If
        End Sub

        Private Sub ResetWorkspaceForAuthenticationRequired()
            CancelActiveThreadSelectionLoad()
            SetSessionAuthenticated(False)
            ResetWorkspaceTransientStateCore(clearModelPicker:=True)
        End Sub

        Private Sub FinalizeAuthenticationRequiredWorkspaceUi()
            UpdateThreadTurnLabels()
            UpdateThreadsStateLabel(0)
            SetTranscriptLoadingState(False)
            RefreshControlStates()
        End Sub

        Private Sub ShowAuthenticationRequiredPromptIfNeeded(showPrompt As Boolean)
            If Not showPrompt OrElse _authRequiredNoticeShown Then
                Return
            End If

            Dim message = "Authentication required. Sign in with ChatGPT or API key to continue."
            AppendSystemMessage(message)
            ShowStatus(message, isError:=True, displayToast:=True)
            _authRequiredNoticeShown = True
        End Sub

        Private Sub NotifyConnectionInitializedUi(isReconnect As Boolean)
            If isReconnect Then
                AppendAndShowSystemMessage("Reconnected and initialized.", displayToast:=True)
            Else
                AppendAndShowSystemMessage("Connected and initialized.", displayToast:=True)
            End If
        End Sub

        Private Sub NotifyRateLimitsUpdatedUi(response As JsonNode)
            _viewModel.SettingsPanel.RateLimitsText = PrettyJson(response)
            MergeRateLimitsPayloadIntoComposerState(response)
            SyncTurnComposerRateLimitBars()
        End Sub

        Private Sub MergeRateLimitsPayloadIntoComposerState(response As JsonNode)
            Dim rootObject = TryCast(response, JsonObject)
            If rootObject Is Nothing Then
                Return
            End If

            MergeRateLimitPayloadObject(rootObject)
            MergeRateLimitPayloadObject(GetPropertyObject(rootObject, "result"))
            MergeRateLimitPayloadObject(GetPropertyObject(rootObject, "params"))
        End Sub

        Private Sub MergeRateLimitPayloadObject(payload As JsonObject)
            If payload Is Nothing Then
                Return
            End If

            Dim byLimitId = GetPropertyObject(payload, "rateLimitsByLimitId")
            If byLimitId IsNot Nothing Then
                For Each entry In byLimitId
                    Dim limitObject = TryCast(entry.Value, JsonObject)
                    If limitObject Is Nothing Then
                        Continue For
                    End If

                    UpsertRateLimitState(limitObject, entry.Key)
                Next
            End If

            Dim singleLimit = GetPropertyObject(payload, "rateLimits")
            If singleLimit IsNot Nothing Then
                UpsertRateLimitState(singleLimit, String.Empty)
            End If
        End Sub

        Private Sub UpsertRateLimitState(limitObject As JsonObject, fallbackLimitId As String)
            If limitObject Is Nothing Then
                Return
            End If

            Dim limitId = GetPropertyString(limitObject, "limitId")
            If String.IsNullOrWhiteSpace(limitId) Then
                limitId = If(fallbackLimitId, String.Empty)
            End If

            limitId = limitId.Trim()
            If String.IsNullOrWhiteSpace(limitId) Then
                Return
            End If

            Dim state As RateLimitLimitState = Nothing
            If Not _rateLimitStatesByLimitId.TryGetValue(limitId, state) Then
                state = New RateLimitLimitState() With {
                    .LimitId = limitId
                }
                _rateLimitStatesByLimitId(limitId) = state
            End If

            state.LimitId = limitId

            Dim limitNameNode As JsonNode = Nothing
            If limitObject.TryGetPropertyValue("limitName", limitNameNode) Then
                state.LimitName = If(limitNameNode Is Nothing, String.Empty, GetPropertyString(limitObject, "limitName").Trim())
            End If

            Dim primaryNode As JsonNode = Nothing
            If limitObject.TryGetPropertyValue("primary", primaryNode) Then
                If primaryNode Is Nothing Then
                    state.Primary = Nothing
                Else
                    Dim primaryObject = TryCast(primaryNode, JsonObject)
                    If primaryObject IsNot Nothing Then
                        state.Primary = MergeRateLimitBucketState(state.Primary, primaryObject)
                    End If
                End If
            End If

            Dim secondaryNode As JsonNode = Nothing
            If limitObject.TryGetPropertyValue("secondary", secondaryNode) Then
                If secondaryNode Is Nothing Then
                    state.Secondary = Nothing
                Else
                    Dim secondaryObject = TryCast(secondaryNode, JsonObject)
                    If secondaryObject IsNot Nothing Then
                        state.Secondary = MergeRateLimitBucketState(state.Secondary, secondaryObject)
                    End If
                End If
            End If
        End Sub

        Private Shared Function MergeRateLimitBucketState(existing As RateLimitBucketState,
                                                          bucketObject As JsonObject) As RateLimitBucketState
            Dim state = If(existing, New RateLimitBucketState())

            Dim usedPercent As Double
            If TryReadJsonDouble(bucketObject, usedPercent, "usedPercent", "used_percent") Then
                state.UsedPercent = ClampPercent(usedPercent)
            End If

            Dim windowDurationMinutes As Integer
            If TryReadJsonInteger(bucketObject, windowDurationMinutes, "windowDurationMins", "window_duration_mins") Then
                state.WindowDurationMins = Math.Max(0, windowDurationMinutes)
            End If

            Dim resetsAtUnix As Long
            If TryReadJsonLong(bucketObject, resetsAtUnix, "resetsAt", "resets_at") Then
                state.ResetsAtUnix = resetsAtUnix
            End If

            Return state
        End Function

        Private Sub SyncTurnComposerRateLimitBars()
            If _viewModel Is Nothing OrElse _viewModel.TurnComposer Is Nothing Then
                Return
            End If

            Dim orderedStates As New List(Of RateLimitLimitState)(_rateLimitStatesByLimitId.Values)
            orderedStates.Sort(AddressOf CompareRateLimitStatesForUi)

            Dim bars As New List(Of TurnComposerRateLimitBarViewModel)()
            For Each state In orderedStates
                If state Is Nothing Then
                    Continue For
                End If

                AppendRateLimitBar(bars, state, "primary", state.Primary)
                AppendRateLimitBar(bars, state, "secondary", state.Secondary)
            Next

            _viewModel.TurnComposer.SetRateLimitBars(bars)
        End Sub

        Private Shared Function CompareRateLimitStatesForUi(left As RateLimitLimitState, right As RateLimitLimitState) As Integer
            Dim leftPriority = If(StringComparer.OrdinalIgnoreCase.Equals(If(left?.LimitId, String.Empty), "codex"), 0, 1)
            Dim rightPriority = If(StringComparer.OrdinalIgnoreCase.Equals(If(right?.LimitId, String.Empty), "codex"), 0, 1)
            If leftPriority <> rightPriority Then
                Return leftPriority.CompareTo(rightPriority)
            End If

            Return StringComparer.OrdinalIgnoreCase.Compare(If(left?.LimitId, String.Empty), If(right?.LimitId, String.Empty))
        End Function

        Private Sub AppendRateLimitBar(target As List(Of TurnComposerRateLimitBarViewModel),
                                       state As RateLimitLimitState,
                                       bucketLabel As String,
                                       bucket As RateLimitBucketState)
            If target Is Nothing OrElse state Is Nothing OrElse bucket Is Nothing OrElse Not bucket.UsedPercent.HasValue Then
                Return
            End If

            Dim usedPercent = ClampPercent(bucket.UsedPercent.Value)
            Dim remainingPercent = ClampPercent(100.0R - usedPercent)
            Dim windowText = If(bucket.WindowDurationMins.HasValue,
                                $"{bucket.WindowDurationMins.Value.ToString(CultureInfo.InvariantCulture)} min",
                                "unknown")
            Dim resetText = FormatRateLimitResetTime(bucket.ResetsAtUnix)
            Dim displayName = ResolveRateLimitDisplayName(state)
            Dim tooltip = $"{displayName} ({bucketLabel}){Environment.NewLine}" &
                          $"Remaining: {remainingPercent.ToString("0.#", CultureInfo.InvariantCulture)}%{Environment.NewLine}" &
                          $"Used: {usedPercent.ToString("0.#", CultureInfo.InvariantCulture)}%{Environment.NewLine}" &
                          $"Window: {windowText}{Environment.NewLine}" &
                          $"Resets: {resetText}"

            target.Add(New TurnComposerRateLimitBarViewModel() With {
                .BarId = $"{state.LimitId}:{bucketLabel}",
                .RemainingPercent = remainingPercent,
                .UsedPercent = usedPercent,
                .TooltipText = tooltip,
                .BarBrush = ResolveRateLimitBarBrush(remainingPercent)
            })
        End Sub

        Private Shared Function ResolveRateLimitDisplayName(state As RateLimitLimitState) As String
            If state Is Nothing Then
                Return "rate-limit"
            End If

            If Not String.IsNullOrWhiteSpace(state.LimitName) Then
                Return state.LimitName
            End If

            If Not String.IsNullOrWhiteSpace(state.LimitId) Then
                Return state.LimitId
            End If

            Return "rate-limit"
        End Function

        Private Function ResolveRateLimitBarBrush(remainingPercent As Double) As Brush
            Dim resourceKey As String
            If remainingPercent <= 20.0R Then
                resourceKey = "DangerBrush"
            ElseIf remainingPercent <= 45.0R Then
                resourceKey = "WarningBrush"
            Else
                resourceKey = "SuccessBrush"
            End If

            Dim resolved = TryCast(TryFindResource(resourceKey), Brush)
            If resolved IsNot Nothing Then
                Return resolved
            End If

            Return TryCast(TryFindResource("AccentBrush"), Brush)
        End Function

        Private Shared Function FormatRateLimitResetTime(resetsAtUnix As Long?) As String
            If Not resetsAtUnix.HasValue Then
                Return "unknown"
            End If

            Try
                Dim unix = resetsAtUnix.Value
                Dim resetTime =
                    If(Math.Abs(unix) > 9_999_999_999L,
                       DateTimeOffset.FromUnixTimeMilliseconds(unix),
                       DateTimeOffset.FromUnixTimeSeconds(unix))
                Return resetTime.LocalDateTime.ToString("yyyy-MM-dd HH:mm:ss")
            Catch
                Return resetsAtUnix.Value.ToString(CultureInfo.InvariantCulture)
            End Try
        End Function

        Private Sub ClearTurnComposerRateLimitStateUi()
            _rateLimitStatesByLimitId.Clear()
            _rateLimitAutoRefreshInProgress = False
            _lastRateLimitAutoRefreshAttemptUtc = DateTimeOffset.MinValue

            If _viewModel IsNot Nothing AndAlso _viewModel.TurnComposer IsNot Nothing Then
                _viewModel.TurnComposer.SetRateLimitBars(Nothing)
                _viewModel.TurnComposer.SetContextUsageIndicator(Nothing, String.Empty)
            End If
        End Sub

        Private Async Function RefreshRateLimitsForComposerIfNeededAsync(force As Boolean,
                                                                         reason As String) As Task
            If _accountService Is Nothing OrElse _viewModel Is Nothing OrElse _viewModel.SessionState Is Nothing Then
                Return
            End If

            If Not _viewModel.SessionState.CanReadRateLimits OrElse Not IsClientRunning() Then
                Return
            End If

            If _rateLimitAutoRefreshInProgress Then
                Return
            End If

            Dim nowUtc = DateTimeOffset.UtcNow
            If Not force Then
                Dim elapsed = nowUtc - _lastRateLimitAutoRefreshAttemptUtc
                If elapsed.TotalSeconds < RateLimitAutoRefreshMinIntervalSeconds Then
                    Return
                End If
            End If

            _rateLimitAutoRefreshInProgress = True
            _lastRateLimitAutoRefreshAttemptUtc = nowUtc
            Try
                Dim response = Await _accountService.ReadRateLimitsAsync(CancellationToken.None).ConfigureAwait(True)
                NotifyRateLimitsUpdatedUi(response)
            Catch ex As Exception
                AppendProtocol("debug",
                               $"rate_limits_auto_refresh_failed reason={If(reason, String.Empty)} error={ex.Message}")
            Finally
                _rateLimitAutoRefreshInProgress = False
            End Try
        End Function

        Private Shared Function ClampPercent(value As Double) As Double
            If Double.IsNaN(value) OrElse Double.IsInfinity(value) Then
                Return 0.0R
            End If

            If value < 0.0R Then
                Return 0.0R
            End If

            If value > 100.0R Then
                Return 100.0R
            End If

            Return value
        End Function

        Private Shared Function TryReadJsonDouble(source As JsonObject,
                                                  ByRef result As Double,
                                                  ParamArray propertyNames() As String) As Boolean
            result = 0.0R
            If source Is Nothing OrElse propertyNames Is Nothing Then
                Return False
            End If

            For Each propertyName In propertyNames
                If String.IsNullOrWhiteSpace(propertyName) Then
                    Continue For
                End If

                Dim node As JsonNode = Nothing
                If Not source.TryGetPropertyValue(propertyName, node) OrElse node Is Nothing Then
                    Continue For
                End If

                Dim jsonValue = TryCast(node, JsonValue)
                If jsonValue Is Nothing Then
                    Continue For
                End If

                Dim doubleValue As Double
                If jsonValue.TryGetValue(Of Double)(doubleValue) Then
                    result = doubleValue
                    Return True
                End If

                Dim integerValue As Integer
                If jsonValue.TryGetValue(Of Integer)(integerValue) Then
                    result = integerValue
                    Return True
                End If

                Dim longValue As Long
                If jsonValue.TryGetValue(Of Long)(longValue) Then
                    result = longValue
                    Return True
                End If

                Dim stringValue As String = Nothing
                If jsonValue.TryGetValue(Of String)(stringValue) AndAlso
                   Double.TryParse(stringValue, NumberStyles.Float, CultureInfo.InvariantCulture, doubleValue) Then
                    result = doubleValue
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function TryReadJsonInteger(source As JsonObject,
                                                   ByRef result As Integer,
                                                   ParamArray propertyNames() As String) As Boolean
            result = 0
            If source Is Nothing OrElse propertyNames Is Nothing Then
                Return False
            End If

            For Each propertyName In propertyNames
                If String.IsNullOrWhiteSpace(propertyName) Then
                    Continue For
                End If

                Dim node As JsonNode = Nothing
                If Not source.TryGetPropertyValue(propertyName, node) OrElse node Is Nothing Then
                    Continue For
                End If

                Dim jsonValue = TryCast(node, JsonValue)
                If jsonValue Is Nothing Then
                    Continue For
                End If

                Dim parsedInteger As Integer
                If jsonValue.TryGetValue(Of Integer)(parsedInteger) Then
                    result = parsedInteger
                    Return True
                End If

                Dim parsedLong As Long
                If jsonValue.TryGetValue(Of Long)(parsedLong) Then
                    If parsedLong >= Integer.MinValue AndAlso parsedLong <= Integer.MaxValue Then
                        result = CInt(parsedLong)
                        Return True
                    End If
                End If

                Dim parsedDouble As Double
                If jsonValue.TryGetValue(Of Double)(parsedDouble) Then
                    result = CInt(Math.Round(parsedDouble, MidpointRounding.AwayFromZero))
                    Return True
                End If

                Dim stringValue As String = Nothing
                If jsonValue.TryGetValue(Of String)(stringValue) AndAlso
                   Integer.TryParse(stringValue, NumberStyles.Integer, CultureInfo.InvariantCulture, parsedInteger) Then
                    result = parsedInteger
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function TryReadJsonLong(source As JsonObject,
                                                ByRef result As Long,
                                                ParamArray propertyNames() As String) As Boolean
            result = 0
            If source Is Nothing OrElse propertyNames Is Nothing Then
                Return False
            End If

            For Each propertyName In propertyNames
                If String.IsNullOrWhiteSpace(propertyName) Then
                    Continue For
                End If

                Dim node As JsonNode = Nothing
                If Not source.TryGetPropertyValue(propertyName, node) OrElse node Is Nothing Then
                    Continue For
                End If

                Dim jsonValue = TryCast(node, JsonValue)
                If jsonValue Is Nothing Then
                    Continue For
                End If

                Dim parsedLong As Long
                If jsonValue.TryGetValue(Of Long)(parsedLong) Then
                    result = parsedLong
                    Return True
                End If

                Dim parsedInteger As Integer
                If jsonValue.TryGetValue(Of Integer)(parsedInteger) Then
                    result = parsedInteger
                    Return True
                End If

                Dim parsedDouble As Double
                If jsonValue.TryGetValue(Of Double)(parsedDouble) Then
                    result = CLng(Math.Round(parsedDouble, MidpointRounding.AwayFromZero))
                    Return True
                End If

                Dim stringValue As String = Nothing
                If jsonValue.TryGetValue(Of String)(stringValue) AndAlso
                   Long.TryParse(stringValue, NumberStyles.Integer, CultureInfo.InvariantCulture, parsedLong) Then
                    result = parsedLong
                    Return True
                End If
            Next

            Return False
        End Function

        Private Sub NotifyApiKeyLoginSubmittedUi()
            AppendAndShowSystemMessage("API key login submitted.", displayToast:=True)
        End Sub

        Private Sub TryOpenChatGptLoginBrowser(authUrl As String)
            Try
                Process.Start(New ProcessStartInfo(authUrl) With {.UseShellExecute = True})
            Catch ex As Exception
                AppendAndShowSystemMessage($"Could not open browser automatically: {ex.Message}",
                                           $"Could not open browser: {ex.Message}",
                                           isError:=True)
            End Try
        End Sub

        Private Sub ApplyChatGptLoginStartedUi(result As ChatGptLoginStartResult)
            SetSessionCurrentLoginId(If(result?.LoginId, String.Empty))
            Dim authUrl = If(result?.AuthUrl, String.Empty)

            If String.IsNullOrWhiteSpace(_viewModel.SessionState.CurrentLoginId) OrElse String.IsNullOrWhiteSpace(authUrl) Then
                Throw New InvalidOperationException("ChatGPT login did not return a valid auth URL.")
            End If

            TryOpenChatGptLoginBrowser(authUrl)
            _viewModel.SettingsPanel.AccountStateText =
                $"Account: waiting for browser sign-in (loginId={_viewModel.SessionState.CurrentLoginId})"
            AppendAndShowSystemMessage("ChatGPT sign-in started. Finish auth in your browser.",
                                       "ChatGPT sign-in started.",
                                       displayToast:=True)
            RefreshControlStates()
        End Sub

        Private Sub NotifyLoginCanceledUi(loginId As String)
            AppendAndShowSystemMessage($"Canceled login {loginId}.")
        End Sub

        Private Sub NotifyLogoutCompletedUi()
            AppendAndShowSystemMessage("Logged out.", displayToast:=True)
        End Sub

        Private Sub BeginLogoutUiTransition()
            _logoutUiTransitionInProgress = True
            RefreshControlStates()
            ShowStatus("Signing out...")
        End Sub

        Private Sub EndLogoutUiTransition()
            _logoutUiTransitionInProgress = False
            RefreshControlStates()
        End Sub

        Private Sub BeginDisconnectUiTransition()
            _disconnectUiTransitionInProgress = True
            RefreshControlStates()
            ShowStatus("Disconnecting...")
        End Sub

        Private Sub EndDisconnectUiTransition()
            _disconnectUiTransitionInProgress = False
            RefreshControlStates()
        End Sub

        Private Sub EnsureAuthenticationRequiredTranscriptTabStripVisible()
            If EnsurePendingNewThreadTranscriptTabActivated() Then
                SetPendingNewThreadFirstPromptSelectionActive(False, clearThreadSelection:=True)
                Return
            End If

            EnsureTranscriptTabSurfaceActivatedForThread(String.Empty)
        End Sub

        Private Sub NotifyExternalTokensAppliedUi()
            AppendAndShowSystemMessage("External ChatGPT auth tokens applied.", displayToast:=True)
        End Sub

        Private Sub UpdateThreadsPanelInteractionState()
            SyncSessionStateViewModel()
            _viewModel.ThreadsPanel.UpdateInteractionState(_viewModel.SessionState,
                                                           _threadsLoading,
                                                           _threadContentLoading)
        End Sub

        Private Sub SyncThreadContentLoadingFromThreadsPanel()
            _threadContentLoading = _viewModel.ThreadsPanel.IsThreadContentLoading
            UpdateThreadsPanelInteractionState()
        End Sub

        Private Function BeginThreadSelectionLoadUiState(threadId As String) As Integer
            Dim loadVersion = _viewModel.ThreadsPanel.BeginThreadSelectionLoad(threadId)
            Dim registration = _threadTranscriptChunkSessionCoordinator.RegisterThreadSelectionLoad(loadVersion,
                                                                                                    threadId,
                                                                                                    "thread_selection_load_begin")
            TraceTranscriptChunkSession("selection_load_begin",
                                        $"loadVersion={registration.UiLoadVersion}; target={registration.ThreadId}; selectionGen={registration.SelectionGenerationId}")
            SyncThreadContentLoadingFromThreadsPanel()
            Return loadVersion
        End Function

        Private Function IsCurrentThreadSelectionLoadUiState(loadVersion As Integer,
                                                             threadId As String) As Boolean
            If Not _viewModel.ThreadsPanel.IsCurrentThreadSelectionLoad(loadVersion, threadId) Then
                Return False
            End If

            Return _threadTranscriptChunkSessionCoordinator.IsCurrentThreadSelectionLoad(loadVersion, threadId)
        End Function

        Private Function TryCompleteThreadSelectionLoadUiState(loadVersion As Integer) As Boolean
            _threadTranscriptChunkSessionCoordinator.CompleteThreadSelectionLoad(loadVersion)
            Dim completed = _viewModel.ThreadsPanel.TryCompleteThreadSelectionLoad(loadVersion)
            If completed Then
                TraceTranscriptChunkSession("selection_load_complete", $"loadVersion={loadVersion}")
                SyncThreadContentLoadingFromThreadsPanel()
            End If

            Return completed
        End Function

        Private Sub ResetThreadSelectionLoadUiState(Optional hideTranscriptLoader As Boolean = False)
            _viewModel.ThreadsPanel.CancelThreadSelectionLoadState()
            _threadTranscriptChunkSessionCoordinator.CancelPendingThreadSelectionLoads("thread_selection_load_ui_reset")
            TraceTranscriptChunkSession("selection_load_reset")
            SyncThreadContentLoadingFromThreadsPanel()
            If hideTranscriptLoader Then
                SetTranscriptLoadingState(False)
            End If
        End Sub

        Private Function CurrentClient() As CodexAppServerClient
            If Not IsClientRunning() Then
                Throw New InvalidOperationException("Not connected to Codex App Server.")
            End If

            Return _client
        End Function

        Private Function EffectiveThreadWorkingDirectory() As String
            Return _viewModel.SettingsPanel.WorkingDir.Trim()
        End Function

        Private Function IsClientRunning() As Boolean
            Return _client IsNot Nothing AndAlso _client.IsRunning
        End Function

        Private Function CanEnterConnectCore() As Boolean
            SyncSessionStateViewModel()
            Return _viewModel.SessionState.CanConnect AndAlso Not IsClientRunning()
        End Function

        Private Sub ApplyConnectedSessionState(client As CodexAppServerClient)
            _client = client
            ApplySessionConnectionEstablished(DateTimeOffset.UtcNow)
            _lastWatchdogWarningUtc = DateTimeOffset.MinValue
            UpdateReconnectCountdownUi()
        End Sub

        Private Sub BeginUserDisconnectSessionTransition()
            SetSessionConnectionExpected(False)
            CancelReconnect()
        End Sub

        Private Function DetachCurrentClient() As CodexAppServerClient
            Dim client = _client
            _client = Nothing
            Return client
        End Function

        Private Function TryBeginManualReconnect(reason As String) As Boolean
            SyncSessionStateViewModel()
            If Not _viewModel.SessionState.CanReconnectNow Then
                ShowStatus("Already connected.")
                Return False
            End If

            CancelReconnect()
            BeginReconnect(reason, force:=True)
            Return True
        End Function

        Private Function TryPrepareReconnectStart(force As Boolean) As Boolean
            SyncSessionStateViewModel()
            Dim session = _viewModel.SessionState

            If Not force AndAlso Not session.ConnectionExpected Then
                Return False
            End If

            If Not force AndAlso Not _viewModel.SettingsPanel.AutoReconnect Then
                SetSessionConnectionExpected(False)
                ShowStatus("Disconnected. Auto-reconnect is disabled.", isError:=True)
                UpdateReconnectCountdownUi()
                Return False
            End If

            If _reconnectInProgress Then
                If force Then
                    CancelReconnect()
                Else
                    Return False
                End If
            End If

            Return Not _reconnectInProgress
        End Function

        Private Sub StartReconnectLoop(reason As String)
            Dim reconnectCts As New CancellationTokenSource()
            _reconnectCts = reconnectCts
            ApplySessionReconnectLoopStartedState()
            ShowStatus($"Auto-reconnect scheduled: {reason}", isError:=True, displayToast:=True)
            AppendSystemMessage($"Auto-reconnect scheduled: {reason}")
            UpdateReconnectCountdownUi()

            FireAndForget(
                _sessionCoordinator.RunReconnectLoopAsync(
                    reason,
                    reconnectCts,
                    Function(token) RunOnUiAsync(Function() ConnectCoreAsync(isReconnect:=True, cancellationToken:=token))))
        End Sub

        Private Async Function ConnectAsync() As Task
            Await ConnectCoreAsync(isReconnect:=False, cancellationToken:=CancellationToken.None)
        End Function

        Private Async Function AutoConnectOnStartupAsync() As Task
            If IsClientRunning() Then
                Return
            End If

            ShowStatus("Auto-connecting to Codex App Server...")
            Await ConnectCoreAsync(isReconnect:=False, cancellationToken:=CancellationToken.None)
        End Function

        Private Async Function ConnectCoreAsync(isReconnect As Boolean,
                                                cancellationToken As CancellationToken) As Task
            If Not CanEnterConnectCore() Then
                Return
            End If

            SaveSettings()

            Dim executable = _viewModel.SettingsPanel.CodexPath.Trim()
            Dim arguments = _viewModel.SettingsPanel.ServerArgs.Trim()
            Dim workingDir = _viewModel.SettingsPanel.WorkingDir.Trim()

            If String.IsNullOrWhiteSpace(executable) Then
                executable = "codex"
            End If

            If String.IsNullOrWhiteSpace(arguments) Then
                arguments = "app-server"
            End If

            If String.IsNullOrWhiteSpace(workingDir) Then
                workingDir = Environment.CurrentDirectory
            End If

            Dim launchExecutable = _connectionService.ResolveWindowsCodexExecutable(executable)
            Dim launchEnvironment = _connectionService.BuildLaunchEnvironment(_viewModel.SettingsPanel.WindowsCodexHome.Trim())

            _viewModel.SettingsPanel.CodexPath = launchExecutable

            Dim client = _connectionService.CreateClient()
            AddHandler client.RawMessage, AddressOf ClientOnRawMessage
            AddHandler client.NotificationReceived, AddressOf ClientOnNotification
            AddHandler client.ServerRequestReceived, AddressOf ClientOnServerRequest
            AddHandler client.Disconnected, AddressOf ClientOnDisconnected

            Dim startupError As Exception = Nothing
            Try
                AppendAndShowSystemMessage($"Starting '{launchExecutable} {arguments}'...")

                Await _connectionService.StartAndInitializeAsync(client,
                                                                 launchExecutable,
                                                                 arguments,
                                                                 workingDir,
                                                                 launchEnvironment,
                                                                 cancellationToken)

                ApplyConnectedSessionState(client)
                NotifyConnectionInitializedUi(isReconnect)

                RefreshControlStates()
                Await EnsureAuthenticatedWithRetryAsync(allowAutoLogin:=True)
                SaveSettings()
            Catch ex As Exception
                startupError = ex
            End Try

            If startupError Is Nothing Then
                Return
            End If

            Try
                RemoveHandler client.RawMessage, AddressOf ClientOnRawMessage
                RemoveHandler client.NotificationReceived, AddressOf ClientOnNotification
                RemoveHandler client.ServerRequestReceived, AddressOf ClientOnServerRequest
                RemoveHandler client.Disconnected, AddressOf ClientOnDisconnected
                Await _connectionService.StopAsync(client, "Connection setup failed.", cancellationToken)
            Catch
            End Try

            Throw startupError
        End Function

        Private Async Function DisconnectAsync() As Task
            BeginDisconnectUiTransition()
            Try
                BeginUserDisconnectSessionTransition()

                Dim client = DetachCurrentClient()
                RefreshControlStates()
                Await Dispatcher.Yield(DispatcherPriority.Render)

                If client IsNot Nothing Then
                    Await DisconnectClientInternalAsync(client,
                                                        "Disconnected by user.",
                                                        CancellationToken.None)
                End If

                ResetDisconnectedUiState("Disconnected.", isError:=False, displayToast:=True)
            Finally
                EndDisconnectUiTransition()
            End Try
        End Function

        Private Async Function DisconnectClientInternalAsync(client As CodexAppServerClient,
                                                             reason As String,
                                                             cancellationToken As CancellationToken) As Task
            If client Is Nothing Then
                Return
            End If

            _disconnecting = True
            Try
                Await Task.Run(
                    Function()
                        Return _connectionService.StopAsync(client, reason, cancellationToken)
                    End Function)
            Finally
                RemoveHandler client.RawMessage, AddressOf ClientOnRawMessage
                RemoveHandler client.NotificationReceived, AddressOf ClientOnNotification
                RemoveHandler client.ServerRequestReceived, AddressOf ClientOnServerRequest
                RemoveHandler client.Disconnected, AddressOf ClientOnDisconnected
                _disconnecting = False
            End Try
        End Function

        Private Sub ResetDisconnectedUiState(message As String,
                                             isError As Boolean,
                                             displayToast As Boolean)
            CancelActiveThreadSelectionLoad()
            ClearPendingUserEchoTracking()
            UpdateConnectionStateTextFromSession()
            SetSessionCurrentLoginId(String.Empty)
            SetSessionAuthenticated(False)
            ClearAuthRequiredNoticeState()
            ResetWorkspaceTransientStateCore(clearModelPicker:=False)
            SetSessionNextReconnectAttempt(Nothing)
            ApplyThreadFiltersAndSort()
            UpdateReconnectCountdownUi()
            UpdateThreadTurnLabels()
            SetTranscriptLoadingState(False)
            RefreshControlStates()
            AppendAndShowSystemMessage(message, isError:=isError, displayToast:=displayToast)
        End Sub

        Private Sub ClientOnRawMessage(direction As String, payload As String)
            RunOnUi(
                Sub()
                    MarkRpcActivity()
                    AppendProtocol(direction, payload)
                End Sub)
        End Sub

        Private Sub ClientOnNotification(methodName As String, paramsNode As JsonNode)
            RunOnUi(
                Sub()
                    HandleNotification(methodName, paramsNode)
                End Sub)
        End Sub

        Private Sub ClientOnServerRequest(request As RpcServerRequest)
            RunOnUi(
                Sub()
                    FireAndForget(HandleServerRequestAsync(request))
                End Sub)
        End Sub

        Private Sub OnTurnWorkflowApprovalResolved(sender As Object,
                                                   e As TurnWorkflowCoordinator.ApprovalResolvedEventArgs)
            If e Is Nothing Then
                Return
            End If

            Dim dispatch = _sessionNotificationCoordinator.DispatchApprovalResolution(e.RequestId, e.Decision)
            ApplyApprovalResolutionDispatchResult(dispatch)
            SyncCurrentTurnFromRuntimeStore(keepExistingWhenRuntimeIsIdle:=True)
            UpdateThreadTurnLabels()
            RefreshThreadRuntimeIndicatorsIfNeeded()
            RefreshControlStates()
        End Sub

        Private Sub ClientOnDisconnected(reason As String)
            RunOnUi(
                Sub()
                    If _disconnecting Then
                        Return
                    End If

                    Dim client = _client
                    _client = Nothing
                    If client IsNot Nothing Then
                        RemoveHandler client.RawMessage, AddressOf ClientOnRawMessage
                        RemoveHandler client.NotificationReceived, AddressOf ClientOnNotification
                        RemoveHandler client.ServerRequestReceived, AddressOf ClientOnServerRequest
                        RemoveHandler client.Disconnected, AddressOf ClientOnDisconnected
                    End If

                    ResetDisconnectedUiState($"Connection closed: {reason}",
                                             isError:=True,
                                             displayToast:=True)
                    BeginReconnect(reason)
                End Sub)
        End Sub

        Private Sub HandleNotification(methodName As String, paramsNode As JsonNode)
            MarkRpcActivity()
            Dim visibleThreadIdBeforeDispatch = GetVisibleThreadId()
            Dim dispatch = _sessionNotificationCoordinator.DispatchNotification(methodName,
                                                                                paramsNode,
                                                                                _notificationRuntimeThreadId,
                                                                                _notificationRuntimeTurnId)
            ApplyNotificationDispatchResult(dispatch)
            If ShouldAppendNotificationTranscriptAuxiliary(dispatch, visibleThreadIdBeforeDispatch) Then
                AppendHandledNotificationMethodEvent(dispatch)
            End If

            SyncCurrentTurnFromRuntimeStore(keepExistingWhenRuntimeIsIdle:=True)
            UpdateThreadTurnLabels()
            RefreshControlStates()
        End Sub

        Private Sub ApplyNotificationDispatchResult(dispatch As SessionNotificationCoordinator.NotificationDispatchResult)
            If dispatch Is Nothing Then
                Return
            End If

            UpdateNotificationRuntimeContextFromDispatch(dispatch)
            Dim visibleThreadIdBeforeDispatch = GetVisibleThreadId()
            ApplyProtocolDispatchMessages(dispatch.ProtocolMessages)
            If ShouldAppendNotificationTranscriptAuxiliary(dispatch, visibleThreadIdBeforeDispatch) Then
                ApplyRuntimeDiagnosticMessages(dispatch.Diagnostics)
            End If

            For Each threadId In dispatch.ThreadIdsToMarkLastActive
                MarkThreadLastActive(threadId)
            Next

            If ShouldAppendNotificationTranscriptAuxiliary(dispatch, visibleThreadIdBeforeDispatch) Then
                For Each message In dispatch.SystemMessages
                    AppendSystemMessage(message)
                Next
            End If

            Dim visibleThreadIdForRuntimeRouting = If(String.IsNullOrWhiteSpace(visibleThreadIdBeforeDispatch),
                                                      GetVisibleThreadId(),
                                                      visibleThreadIdBeforeDispatch)
            Dim visibleRuntimeItemsRendered = 0
            Dim didMutateVisibleTranscript = False

            For Each item In dispatch.RuntimeItems
                If item Is Nothing Then
                    Continue For
                End If

                Dim resolvedThreadId As String = Nothing
                Dim resolvedTurnId As String = Nothing
                If Not TryResolveRuntimeEventUiScope(item.ThreadId,
                                                    item.TurnId,
                                                    _notificationRuntimeThreadId,
                                                    _notificationRuntimeTurnId,
                                                    resolvedThreadId,
                                                    resolvedTurnId) Then
                    Continue For
                End If

                If Not HandleRuntimeEventForThreadVisibility(resolvedThreadId,
                                                             resolvedTurnId,
                                                             visibleThreadIdForRuntimeRouting,
                                                             dispatch.MethodName) Then
                    Continue For
                End If

                ' Avoid rendering ambiguous item snapshots into the visible thread when the item itself
                ' still lacks concrete scope; wait for a concrete-scoped update or rebuild replay.
                If Not StringComparer.Ordinal.Equals(If(item.ThreadId, String.Empty).Trim(), resolvedThreadId) OrElse
                   IsInferredRuntimeScopeId(item.ThreadId) Then
                    Continue For
                End If

                If item.IsCompleted AndAlso
                   StringComparer.OrdinalIgnoreCase.Equals(If(item.ItemType, String.Empty), "contextCompaction") Then
                    TraceContextUsageDebug("context_compaction_completed",
                                           $"thread={resolvedThreadId}; turn={resolvedTurnId}; triggering_thread_read_refresh=True")
                    FireAndForget(RefreshSelectedThreadTokenUsageFromServerAsync(resolvedThreadId))
                End If

                RenderItem(item)
                visibleRuntimeItemsRendered += 1
                didMutateVisibleTranscript = True
            Next

            For Each lifecycleMessage In dispatch.TurnLifecycleMessages
                If lifecycleMessage Is Nothing Then
                    Continue For
                End If

                Dim resolvedThreadId As String = Nothing
                Dim resolvedTurnId As String = Nothing
                If Not TryResolveRuntimeEventUiScope(lifecycleMessage.ThreadId,
                                                    lifecycleMessage.TurnId,
                                                    _notificationRuntimeThreadId,
                                                    _notificationRuntimeTurnId,
                                                    resolvedThreadId,
                                                    resolvedTurnId) Then
                    Continue For
                End If

                If IsTurnLifecycleCompletionStatus(lifecycleMessage.Status) Then
                    PlayTurnDoneSoundIfEnabled(resolvedThreadId, resolvedTurnId)
                End If

                If Not HandleRuntimeEventForThreadVisibility(resolvedThreadId,
                                                             resolvedTurnId,
                                                             visibleThreadIdForRuntimeRouting,
                                                             dispatch.MethodName) Then
                    Continue For
                End If

                AppendTurnLifecycleMarker(resolvedThreadId,
                                          resolvedTurnId,
                                          lifecycleMessage.Status)
                didMutateVisibleTranscript = True
            Next

            For Each metadataMessage In dispatch.TurnMetadataMessages
                If metadataMessage Is Nothing Then
                    Continue For
                End If

                Dim resolvedThreadId As String = Nothing
                Dim resolvedTurnId As String = Nothing
                If Not TryResolveRuntimeEventUiScope(metadataMessage.ThreadId,
                                                    metadataMessage.TurnId,
                                                    _notificationRuntimeThreadId,
                                                    _notificationRuntimeTurnId,
                                                    resolvedThreadId,
                                                    resolvedTurnId) Then
                    Continue For
                End If

                If Not HandleRuntimeEventForThreadVisibility(resolvedThreadId,
                                                             resolvedTurnId,
                                                             visibleThreadIdForRuntimeRouting,
                                                             dispatch.MethodName) Then
                    Continue For
                End If

                UpsertTurnMetadata(resolvedThreadId,
                                   resolvedTurnId,
                                   metadataMessage.Kind,
                                   metadataMessage.SummaryText)
                didMutateVisibleTranscript = True
            Next

            For Each tokenUsageMessage In dispatch.TokenUsageMessages
                If tokenUsageMessage Is Nothing Then
                    Continue For
                End If

                Dim resolvedThreadId As String = Nothing
                Dim resolvedTurnId As String = Nothing
                If Not TryResolveRuntimeEventUiScope(tokenUsageMessage.ThreadId,
                                                    tokenUsageMessage.TurnId,
                                                    _notificationRuntimeThreadId,
                                                    _notificationRuntimeTurnId,
                                                    resolvedThreadId,
                                                    resolvedTurnId) Then
                    TraceContextUsageDebug("notification_token_usage_scope_miss",
                                           $"rawThread={If(tokenUsageMessage.ThreadId, String.Empty).Trim()}; rawTurn={If(tokenUsageMessage.TurnId, String.Empty).Trim()}; notificationThread={If(_notificationRuntimeThreadId, String.Empty).Trim()}; notificationTurn={If(_notificationRuntimeTurnId, String.Empty).Trim()}")
                    Continue For
                End If

                If Not HandleRuntimeEventForThreadVisibility(resolvedThreadId,
                                                             resolvedTurnId,
                                                             visibleThreadIdForRuntimeRouting,
                                                             dispatch.MethodName) Then
                    TraceContextUsageDebug("notification_token_usage_not_visible",
                                           $"thread={resolvedThreadId}; turn={resolvedTurnId}; visibleThread={If(visibleThreadIdForRuntimeRouting, String.Empty).Trim()}; method={If(dispatch.MethodName, String.Empty)}")
                    Continue For
                End If

                TraceContextUsageDebug("notification_token_usage_apply",
                                       $"thread={resolvedThreadId}; turn={resolvedTurnId}; keys={DescribeJsonObjectKeys(tokenUsageMessage.TokenUsage)}")
                UpdateTokenUsageWidget(resolvedThreadId,
                                       resolvedTurnId,
                                       tokenUsageMessage.TokenUsage)
            Next

            For Each loginId In dispatch.LoginIdsToClear
                If StringComparer.Ordinal.Equals(loginId, _viewModel.SessionState.CurrentLoginId) Then
                    SetSessionCurrentLoginId(String.Empty)
                End If
            Next

            If dispatch.ShouldRefreshAuthentication Then
                FireAndForget(RunUiActionAsync(AddressOf RefreshAuthenticationGateAsync))
            End If

            For Each rateLimitsPayload In dispatch.RateLimitPayloads
                NotifyRateLimitsUpdatedUi(rateLimitsPayload)
            Next

            If (dispatch.ShouldScrollTranscriptToBottom AndAlso didMutateVisibleTranscript) OrElse
               visibleRuntimeItemsRendered > 0 Then
                ScrollTranscriptToBottom(reason:=TranscriptScrollRequestReason.RuntimeStream)
            End If

            TryDispatchDeferredTranscriptChunkPrependFromRuntimeUpdate()
        End Sub

        Private Shared Function IsTurnLifecycleCompletionStatus(status As String) As Boolean
            Select Case If(status, String.Empty).Trim().ToLowerInvariant()
                Case "", "completed", "complete", "succeeded", "success", "ok", "done", "failed", "error", "canceled", "cancelled", "interrupted", "aborted", "stopped"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Sub UpdateNotificationRuntimeContextFromDispatch(dispatch As SessionNotificationCoordinator.NotificationDispatchResult)
            If dispatch Is Nothing Then
                Return
            End If

            Dim runtimeThreadId = If(_notificationRuntimeThreadId, String.Empty).Trim()
            Dim runtimeTurnId = If(_notificationRuntimeTurnId, String.Empty).Trim()

            If dispatch.ThreadObject IsNot Nothing Then
                Dim threadIdFromObject = GetPropertyString(dispatch.ThreadObject, "id")
                If Not String.IsNullOrWhiteSpace(threadIdFromObject) Then
                    runtimeThreadId = threadIdFromObject.Trim()
                End If
            End If

            If dispatch.CurrentThreadChanged Then
                runtimeThreadId = If(dispatch.CurrentThreadId, String.Empty).Trim()
            ElseIf String.IsNullOrWhiteSpace(runtimeThreadId) AndAlso
                   Not String.IsNullOrWhiteSpace(dispatch.CurrentThreadId) Then
                runtimeThreadId = dispatch.CurrentThreadId.Trim()
            End If

            If dispatch.CurrentTurnChanged Then
                runtimeTurnId = If(dispatch.CurrentTurnId, String.Empty).Trim()
            ElseIf String.IsNullOrWhiteSpace(runtimeTurnId) AndAlso
                   Not String.IsNullOrWhiteSpace(dispatch.CurrentTurnId) Then
                runtimeTurnId = dispatch.CurrentTurnId.Trim()
            End If

            ' If a thread context changed and no explicit turn is active, clear the stale turn fallback.
            If dispatch.CurrentThreadChanged AndAlso String.IsNullOrWhiteSpace(If(dispatch.CurrentTurnId, String.Empty).Trim()) Then
                runtimeTurnId = String.Empty
            End If

            _notificationRuntimeThreadId = runtimeThreadId
            _notificationRuntimeTurnId = runtimeTurnId
        End Sub

        Private Sub ApplyServerRequestDispatchResult(dispatch As SessionNotificationCoordinator.ServerRequestDispatchResult)
            If dispatch Is Nothing Then
                Return
            End If

            ApplyProtocolDispatchMessages(dispatch.ProtocolMessages)

            Dim visibleThreadIdForRuntimeRouting = GetVisibleThreadId()
            Dim visibleRuntimeItemsRendered = 0
            Dim anyRuntimeItemsObserved = False
            For Each item In dispatch.RuntimeItems
                If item Is Nothing Then
                    Continue For
                End If
                anyRuntimeItemsObserved = True

                Dim resolvedThreadId As String = Nothing
                Dim resolvedTurnId As String = Nothing
                If Not TryResolveRuntimeEventUiScope(item.ThreadId,
                                                    item.TurnId,
                                                    Nothing,
                                                    Nothing,
                                                    resolvedThreadId,
                                                    resolvedTurnId) Then
                    Continue For
                End If

                If Not HandleRuntimeEventForThreadVisibility(resolvedThreadId,
                                                             resolvedTurnId,
                                                             visibleThreadIdForRuntimeRouting,
                                                             dispatch.MethodName) Then
                    Continue For
                End If

                If Not StringComparer.Ordinal.Equals(If(item.ThreadId, String.Empty).Trim(), resolvedThreadId) OrElse
                   IsInferredRuntimeScopeId(item.ThreadId) Then
                    Continue For
                End If

                RenderItem(item)
                visibleRuntimeItemsRendered += 1
            Next

            If visibleRuntimeItemsRendered > 0 OrElse Not anyRuntimeItemsObserved Then
                ApplyRuntimeDiagnosticMessages(dispatch.Diagnostics)
            End If

            If visibleRuntimeItemsRendered > 0 Then
                ScrollTranscriptToBottom(reason:=TranscriptScrollRequestReason.RuntimeStream)
            End If

            RefreshThreadRuntimeIndicatorsIfNeeded()
            TryDispatchDeferredTranscriptChunkPrependFromRuntimeUpdate()
        End Sub

        Private Sub ApplyApprovalResolutionDispatchResult(dispatch As SessionNotificationCoordinator.ApprovalResolutionDispatchResult)
            If dispatch Is Nothing Then
                Return
            End If

            ApplyProtocolDispatchMessages(dispatch.ProtocolMessages)

            Dim visibleThreadIdForRuntimeRouting = GetVisibleThreadId()
            Dim visibleRuntimeItemsRendered = 0
            Dim anyRuntimeItemsObserved = False
            For Each item In dispatch.RuntimeItems
                If item Is Nothing Then
                    Continue For
                End If
                anyRuntimeItemsObserved = True

                Dim resolvedThreadId As String = Nothing
                Dim resolvedTurnId As String = Nothing
                If Not TryResolveRuntimeEventUiScope(item.ThreadId,
                                                    item.TurnId,
                                                    Nothing,
                                                    Nothing,
                                                    resolvedThreadId,
                                                    resolvedTurnId) Then
                    Continue For
                End If

                If Not HandleRuntimeEventForThreadVisibility(resolvedThreadId,
                                                             resolvedTurnId,
                                                             visibleThreadIdForRuntimeRouting,
                                                             "approval_resolution") Then
                    Continue For
                End If

                If Not StringComparer.Ordinal.Equals(If(item.ThreadId, String.Empty).Trim(), resolvedThreadId) OrElse
                   IsInferredRuntimeScopeId(item.ThreadId) Then
                    Continue For
                End If

                RenderItem(item)
                visibleRuntimeItemsRendered += 1
            Next

            If visibleRuntimeItemsRendered > 0 OrElse Not anyRuntimeItemsObserved Then
                ApplyRuntimeDiagnosticMessages(dispatch.Diagnostics)
            End If

            If visibleRuntimeItemsRendered > 0 Then
                ScrollTranscriptToBottom(reason:=TranscriptScrollRequestReason.RuntimeStream)
            End If

            RefreshThreadRuntimeIndicatorsIfNeeded()
            TryDispatchDeferredTranscriptChunkPrependFromRuntimeUpdate()
        End Sub

        Private Function TryResolveRuntimeEventUiScope(threadId As String,
                                                       turnId As String,
                                                       fallbackThreadId As String,
                                                       fallbackTurnId As String,
                                                       ByRef resolvedThreadId As String,
                                                       ByRef resolvedTurnId As String) As Boolean
            resolvedThreadId = If(threadId, String.Empty).Trim()
            resolvedTurnId = If(turnId, String.Empty).Trim()

            Dim normalizedFallbackThreadId = If(fallbackThreadId, String.Empty).Trim()
            Dim normalizedFallbackTurnId = If(fallbackTurnId, String.Empty).Trim()

            If IsInferredRuntimeScopeId(resolvedTurnId) Then
                resolvedTurnId = String.Empty
            End If
            If String.IsNullOrWhiteSpace(resolvedTurnId) AndAlso Not String.IsNullOrWhiteSpace(normalizedFallbackTurnId) Then
                resolvedTurnId = normalizedFallbackTurnId
            End If

            If IsInferredRuntimeScopeId(resolvedThreadId) Then
                resolvedThreadId = String.Empty
            End If

            If String.IsNullOrWhiteSpace(resolvedThreadId) AndAlso
               Not String.IsNullOrWhiteSpace(resolvedTurnId) Then
                Dim threadIdFromTurn As String = Nothing
                If TryResolveThreadIdFromRuntimeTurn(resolvedTurnId, threadIdFromTurn) Then
                    resolvedThreadId = threadIdFromTurn
                End If
            End If

            If String.IsNullOrWhiteSpace(resolvedThreadId) AndAlso
               Not String.IsNullOrWhiteSpace(normalizedFallbackThreadId) Then
                resolvedThreadId = normalizedFallbackThreadId
            End If

            If IsInferredRuntimeScopeId(resolvedThreadId) Then
                resolvedThreadId = String.Empty
            End If

            Return Not String.IsNullOrWhiteSpace(resolvedThreadId)
        End Function

        Private Function TryResolveThreadIdFromRuntimeTurn(turnId As String, ByRef resolvedThreadId As String) As Boolean
            resolvedThreadId = String.Empty

            Dim normalizedTurnId = If(turnId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedTurnId) OrElse IsInferredRuntimeScopeId(normalizedTurnId) Then
                Return False
            End If

            Dim runtimeStore = _sessionNotificationCoordinator.RuntimeStore
            If runtimeStore Is Nothing OrElse runtimeStore.TurnsById Is Nothing Then
                Return False
            End If

            For Each pair In runtimeStore.TurnsById
                Dim turn = pair.Value
                If turn Is Nothing Then
                    Continue For
                End If

                If Not StringComparer.Ordinal.Equals(If(turn.TurnId, String.Empty).Trim(), normalizedTurnId) Then
                    Continue For
                End If

                Dim candidateThreadId = If(turn.ThreadId, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(candidateThreadId) OrElse IsInferredRuntimeScopeId(candidateThreadId) Then
                    Continue For
                End If

                resolvedThreadId = candidateThreadId
                Return True
            Next

            Return False
        End Function

        Private Shared Function IsInferredRuntimeScopeId(value As String) As Boolean
            Dim normalized = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalized) Then
                Return False
            End If

            Return normalized.StartsWith("__inferred_", StringComparison.Ordinal)
        End Function

        Private Function HandleRuntimeEventForThreadVisibility(threadId As String,
                                                               turnId As String,
                                                               visibleThreadId As String,
                                                               methodName As String) As Boolean
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            Dim normalizedTurnId = If(turnId, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse IsInferredRuntimeScopeId(normalizedThreadId) Then
                Return False
            End If

            If IsInferredRuntimeScopeId(normalizedTurnId) Then
                normalizedTurnId = String.Empty
            End If

            Dim isVisible = IsVisibleThread(normalizedThreadId, visibleThreadId)
            If isVisible Then
                UpdateThreadLiveSessionRuntimeActivity(normalizedThreadId, normalizedTurnId)
                If _sessionNotificationCoordinator.RuntimeStore.HasActiveTurn(normalizedThreadId) Then
                    TryMarkThreadOverlayTurn(normalizedThreadId, normalizedTurnId, allowCompletedTurnFallback:=False)
                End If
                _threadLiveSessionRegistry.SetPendingRebuild(normalizedThreadId, False)
                _threadLiveSessionRegistry.MarkBound(normalizedThreadId, GetVisibleTurnId())
                Return True
            End If

            MarkThreadLiveSessionDirty(normalizedThreadId, normalizedTurnId, methodName)
            Return False
        End Function

        Private Sub MarkThreadLiveSessionDirty(threadId As String, turnId As String, methodName As String)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            UpdateThreadLiveSessionRuntimeActivity(normalizedThreadId, turnId)
            TryMarkThreadOverlayTurn(normalizedThreadId, turnId, allowCompletedTurnFallback:=True)
            _threadLiveSessionRegistry.SetPendingRebuild(normalizedThreadId, True)
        End Sub

        Private Sub UpdateThreadLiveSessionRuntimeActivity(threadId As String, turnId As String)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            Dim normalizedTurnId = If(turnId, String.Empty).Trim()
            Dim runtimeStore = _sessionNotificationCoordinator.RuntimeStore

            Dim hasActiveTurn = runtimeStore.HasActiveTurn(normalizedThreadId)
            Dim activeTurnId = runtimeStore.GetActiveTurnId(normalizedThreadId, normalizedTurnId)
            Dim relevantTurnId = activeTurnId
            If String.IsNullOrWhiteSpace(relevantTurnId) Then
                relevantTurnId = normalizedTurnId
            End If
            If String.IsNullOrWhiteSpace(relevantTurnId) Then
                relevantTurnId = runtimeStore.GetLatestTurnId(normalizedThreadId)
            End If

            _threadLiveSessionRegistry.MarkRuntimeActivity(normalizedThreadId,
                                                           relevantTurnId,
                                                           hasActiveTurn)
        End Sub

        Private Sub TryMarkThreadOverlayTurn(threadId As String,
                                             turnId As String,
                                             allowCompletedTurnFallback As Boolean)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            Dim resolvedTurnId = ResolveOverlayTurnIdForRuntimeEvent(normalizedThreadId,
                                                                     turnId,
                                                                     allowCompletedTurnFallback)
            If String.IsNullOrWhiteSpace(resolvedTurnId) Then
                Return
            End If

            _threadLiveSessionRegistry.MarkOverlayTurn(normalizedThreadId, resolvedTurnId)
        End Sub

        Private Function ResolveOverlayTurnIdForRuntimeEvent(threadId As String,
                                                             turnId As String,
                                                             allowCompletedTurnFallback As Boolean) As String
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return String.Empty
            End If

            Dim normalizedTurnId = If(turnId, String.Empty).Trim()
            Dim runtimeStore = _sessionNotificationCoordinator.RuntimeStore

            Dim activeTurnId = runtimeStore.GetActiveTurnId(normalizedThreadId, normalizedTurnId)
            If Not String.IsNullOrWhiteSpace(activeTurnId) Then
                Return activeTurnId
            End If

            If allowCompletedTurnFallback Then
                If Not String.IsNullOrWhiteSpace(normalizedTurnId) Then
                    Return normalizedTurnId
                End If

                Return runtimeStore.GetLatestTurnId(normalizedThreadId)
            End If

            Return String.Empty
        End Function

        Private Function IsVisibleThread(threadId As String, Optional visibleThreadId As String = Nothing) As Boolean
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return False
            End If

            Dim baselineVisibleThreadId = If(visibleThreadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(baselineVisibleThreadId) Then
                baselineVisibleThreadId = GetVisibleThreadId()
            End If

            If String.IsNullOrWhiteSpace(baselineVisibleThreadId) Then
                Return False
            End If

            Return StringComparer.Ordinal.Equals(normalizedThreadId, baselineVisibleThreadId)
        End Function

        Private Function ShouldAppendNotificationTranscriptAuxiliary(dispatch As SessionNotificationCoordinator.NotificationDispatchResult,
                                                                    baselineVisibleThreadId As String) As Boolean
            If dispatch Is Nothing Then
                Return False
            End If

            Dim normalizedVisibleThreadId = If(baselineVisibleThreadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedVisibleThreadId) Then
                Return True
            End If

            Dim sawScopedPayload = False
            Dim sawVisibleScopedPayload = False

            Dim resolvedThreadId As String = Nothing
            Dim resolvedTurnId As String = Nothing

            For Each item In dispatch.RuntimeItems
                If item Is Nothing Then
                    Continue For
                End If

                If Not TryResolveRuntimeEventUiScope(item.ThreadId,
                                                    item.TurnId,
                                                    _notificationRuntimeThreadId,
                                                    _notificationRuntimeTurnId,
                                                    resolvedThreadId,
                                                    resolvedTurnId) Then
                    Continue For
                End If

                sawScopedPayload = True
                If IsVisibleThread(resolvedThreadId, normalizedVisibleThreadId) Then
                    sawVisibleScopedPayload = True
                    Exit For
                End If
            Next

            If Not sawVisibleScopedPayload Then
                For Each lifecycleMessage In dispatch.TurnLifecycleMessages
                    If lifecycleMessage Is Nothing Then
                        Continue For
                    End If

                    If Not TryResolveRuntimeEventUiScope(lifecycleMessage.ThreadId,
                                                        lifecycleMessage.TurnId,
                                                        _notificationRuntimeThreadId,
                                                        _notificationRuntimeTurnId,
                                                        resolvedThreadId,
                                                        resolvedTurnId) Then
                        Continue For
                    End If

                    sawScopedPayload = True
                    If IsVisibleThread(resolvedThreadId, normalizedVisibleThreadId) Then
                        sawVisibleScopedPayload = True
                        Exit For
                    End If
                Next
            End If

            If Not sawVisibleScopedPayload Then
                For Each metadataMessage In dispatch.TurnMetadataMessages
                    If metadataMessage Is Nothing Then
                        Continue For
                    End If

                    If Not TryResolveRuntimeEventUiScope(metadataMessage.ThreadId,
                                                        metadataMessage.TurnId,
                                                        _notificationRuntimeThreadId,
                                                        _notificationRuntimeTurnId,
                                                        resolvedThreadId,
                                                        resolvedTurnId) Then
                        Continue For
                    End If

                    sawScopedPayload = True
                    If IsVisibleThread(resolvedThreadId, normalizedVisibleThreadId) Then
                        sawVisibleScopedPayload = True
                        Exit For
                    End If
                Next
            End If

            If Not sawVisibleScopedPayload Then
                For Each tokenUsageMessage In dispatch.TokenUsageMessages
                    If tokenUsageMessage Is Nothing Then
                        Continue For
                    End If

                    If Not TryResolveRuntimeEventUiScope(tokenUsageMessage.ThreadId,
                                                        tokenUsageMessage.TurnId,
                                                        _notificationRuntimeThreadId,
                                                        _notificationRuntimeTurnId,
                                                        resolvedThreadId,
                                                        resolvedTurnId) Then
                        Continue For
                    End If

                    sawScopedPayload = True
                    If IsVisibleThread(resolvedThreadId, normalizedVisibleThreadId) Then
                        sawVisibleScopedPayload = True
                        Exit For
                    End If
                Next
            End If

            If sawScopedPayload AndAlso Not sawVisibleScopedPayload Then
                Return False
            End If

            Dim scopedThreadId = If(_notificationRuntimeThreadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(scopedThreadId) Then
                scopedThreadId = If(dispatch.CurrentThreadId, String.Empty).Trim()
            End If

            ' Keep global/unscoped notifications visible (for example auth/rate-limit/system notices).
            If String.IsNullOrWhiteSpace(scopedThreadId) Then
                Return True
            End If

            Return StringComparer.Ordinal.Equals(scopedThreadId, normalizedVisibleThreadId)
        End Function

        Private Sub ApplyProtocolDispatchMessages(messages As List(Of SessionNotificationCoordinator.ProtocolDispatchMessage))
            If messages Is Nothing Then
                Return
            End If

            For Each message In messages
                If message Is Nothing Then
                    Continue For
                End If

                AppendProtocol(message.Direction, message.Payload)
            Next
        End Sub

        Private Sub ApplyRuntimeDiagnosticMessages(messages As List(Of String))
            If messages Is Nothing Then
                Return
            End If

            For Each message In messages
                If String.IsNullOrWhiteSpace(message) Then
                    Continue For
                End If

                AppendRuntimeDiagnosticEvent(message)
            Next
        End Sub

        Private Sub AppendHandledNotificationMethodEvent(dispatch As SessionNotificationCoordinator.NotificationDispatchResult)
            If dispatch Is Nothing Then
                Return
            End If

            Dim normalizedMethod = If(dispatch.MethodName, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedMethod) Then
                Return
            End If

            If NotificationDispatchContainsUnhandledMethodDiagnostic(dispatch) Then
                Return
            End If

            If NotificationDispatchContainsSkippedMethodDiagnostic(dispatch, normalizedMethod) Then
                Return
            End If

            AppendRuntimeDiagnosticEvent($"Handled notification method: {normalizedMethod}")
        End Sub

        Private Shared Function NotificationDispatchContainsUnhandledMethodDiagnostic(dispatch As SessionNotificationCoordinator.NotificationDispatchResult) As Boolean
            If dispatch Is Nothing OrElse dispatch.Diagnostics Is Nothing Then
                Return False
            End If

            For Each diagnostic In dispatch.Diagnostics
                If String.IsNullOrWhiteSpace(diagnostic) Then
                    Continue For
                End If

                If diagnostic.StartsWith("Unhandled notification method:", StringComparison.Ordinal) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function NotificationDispatchContainsSkippedMethodDiagnostic(dispatch As SessionNotificationCoordinator.NotificationDispatchResult,
                                                                                    methodName As String) As Boolean
            If dispatch Is Nothing OrElse dispatch.Diagnostics Is Nothing Then
                Return False
            End If

            Dim normalizedMethod = If(methodName, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedMethod) Then
                Return False
            End If

            Dim skipPrefix = normalizedMethod & ":"
            For Each diagnostic In dispatch.Diagnostics
                If String.IsNullOrWhiteSpace(diagnostic) Then
                    Continue For
                End If

                If diagnostic.StartsWith(skipPrefix, StringComparison.Ordinal) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Function ReadPersistedApiKey() As String
            If _settings Is Nothing OrElse String.IsNullOrWhiteSpace(_settings.EncryptedApiKey) Then
                Return String.Empty
            End If

            Try
                Dim encryptedBytes = Convert.FromBase64String(_settings.EncryptedApiKey)
                Dim plainBytes = ProtectedData.Unprotect(encryptedBytes, Nothing, DataProtectionScope.CurrentUser)
                Return Encoding.UTF8.GetString(plainBytes)
            Catch
                Return String.Empty
            End Try
        End Function

        Private Sub PersistApiKeyIfNeeded(apiKey As String)
            If _settings Is Nothing Then
                _settings = New AppSettings()
            End If

            If _viewModel.SettingsPanel.RememberApiKey AndAlso Not String.IsNullOrWhiteSpace(apiKey) Then
                Dim plainBytes = Encoding.UTF8.GetBytes(apiKey)
                Dim encryptedBytes = ProtectedData.Protect(plainBytes, Nothing, DataProtectionScope.CurrentUser)
                _settings.EncryptedApiKey = Convert.ToBase64String(encryptedBytes)
            Else
                _settings.EncryptedApiKey = String.Empty
            End If

            SaveSettings()
        End Sub

        Private Async Function TryAutoLoginApiKeyAsync() As Task
            If Not _viewModel.SettingsPanel.AutoLoginApiKey Then
                Return
            End If

            Dim storedApiKey = ReadPersistedApiKey()
            If String.IsNullOrWhiteSpace(storedApiKey) Then
                Return
            End If

            Try
                Await _accountService.StartApiKeyLoginAsync(storedApiKey, CancellationToken.None)
                AppendSystemMessage("Auto-login with saved API key completed.")
                ShowStatus("Auto-login with saved API key completed.")
            Catch ex As Exception
                AppendSystemMessage($"Auto-login with saved API key failed: {ex.Message}")
                ShowStatus($"Auto-login failed: {ex.Message}", isError:=True)
            End Try
        End Function

        Private Async Function EnsureAuthenticatedAndInitializeWorkspaceAsync(allowAutoLogin As Boolean,
                                                                              showAuthPrompt As Boolean) As Task(Of Boolean)
            SyncSessionStateViewModel()
            Dim wasAuthenticated = _viewModel.SessionState.IsAuthenticated
            Dim hasAccount = Await RefreshAccountAsync(False)

            If Not hasAccount Then
                hasAccount = Await RefreshAccountAsync(True)
            End If

            If Not hasAccount AndAlso allowAutoLogin Then
                Await TryAutoLoginApiKeyAsync()
                hasAccount = Await RefreshAccountAsync(False)
                If Not hasAccount Then
                    hasAccount = Await RefreshAccountAsync(True)
                End If
            End If

            If hasAccount Then
                ClearAuthRequiredNoticeState()
                Dim shouldInitializeWorkspace = ShouldInitializeWorkspaceAfterAuthentication(wasAuthenticated)
                Dim hasCachedRateLimits = _rateLimitStatesByLimitId.Count > 0
                Await RefreshRateLimitsForComposerIfNeededAsync(force:=(shouldInitializeWorkspace OrElse Not hasCachedRateLimits),
                                                                 reason:="auth_ready")

                If shouldInitializeWorkspace Then
                    Await InitializeWorkspaceAfterAuthenticationAsync()
                End If

                ShowThreadsSidebarTab()
                Return True
            End If

            ApplyAuthenticationRequiredState(showPrompt:=showAuthPrompt)
            Return False
        End Function

        Private Async Function EnsureAuthenticatedWithRetryAsync(allowAutoLogin As Boolean) As Task(Of Boolean)
            Return Await _sessionCoordinator.RunAuthenticationRetryAsync(
                allowAutoLogin,
                Function(useAutoLogin, showAuthPrompt)
                    Return EnsureAuthenticatedAndInitializeWorkspaceAsync(useAutoLogin, showAuthPrompt)
                End Function)
        End Function

        Private Async Function RefreshAuthenticationGateAsync() As Task
            Await _sessionCoordinator.RunAuthenticationRetryAsync(
                allowAutoLogin:=False,
                ensureAttemptAsync:=Function(useAutoLogin, showAuthPrompt)
                                        Return EnsureAuthenticatedAndInitializeWorkspaceAsync(useAutoLogin, showAuthPrompt)
                                    End Function)
        End Function

        Private Async Function InitializeWorkspaceAfterAuthenticationAsync() As Task
            If Not CanBeginWorkspaceBootstrapAfterAuthentication() Then
                Return
            End If

            Try
                Await RefreshModelsAsync()
                Await RefreshThreadsAsync()
                InitializeStartupDraftNewThreadUi()
                ShowStatus("Connected and authenticated.")
            Finally
                EndWorkspaceBootstrapAfterAuthentication()
            End Try
        End Function

        Private Sub ApplyAuthenticationRequiredState(Optional showPrompt As Boolean = False)
            ResetWorkspaceForAuthenticationRequired()
            FinalizeAuthenticationRequiredWorkspaceUi()
            EnsureAuthenticationRequiredTranscriptTabStripVisible()
            ShowAuthenticationRequiredPromptIfNeeded(showPrompt)

            ShowControlCenterTab()
        End Sub

        Private Async Function RefreshAccountAsync(refreshToken As Boolean) As Task(Of Boolean)
            Return Await _sessionCoordinator.ReadAccountAndApplyAsync(
                refreshToken,
                Function(useRefreshToken, token)
                    Return _accountService.ReadAccountAsync(useRefreshToken, token)
                End Function,
                Function(result)
                    Return ApplyAccountReadResultUi(result)
                End Function)
        End Function

        Private Async Function ReadRateLimitsAsync() As Task
            Await _sessionCoordinator.ReadRateLimitsAndApplyAsync(
                Function(token)
                    Return _accountService.ReadRateLimitsAsync(token)
                End Function,
                Sub(response)
                    NotifyRateLimitsUpdatedUi(response)
                End Sub)
        End Function

        Private Async Function LoginApiKeyAsync() As Task
            Dim apiKey = _viewModel.SettingsPanel.ApiKey.Trim()
            If String.IsNullOrWhiteSpace(apiKey) Then
                Throw New InvalidOperationException("Enter an OpenAI API key first.")
            End If

            Await _accountService.StartApiKeyLoginAsync(apiKey, CancellationToken.None)
            PersistApiKeyIfNeeded(apiKey)
            NotifyApiKeyLoginSubmittedUi()
            Await EnsureAuthenticatedWithRetryAsync(allowAutoLogin:=False)
        End Function

        Private Async Function LoginChatGptAsync() As Task
            Dim result = Await _accountService.StartChatGptLoginAsync(CancellationToken.None)
            ApplyChatGptLoginStartedUi(result)
        End Function

        Private Async Function CancelLoginAsync() As Task
            If String.IsNullOrWhiteSpace(_viewModel.SessionState.CurrentLoginId) Then
                Throw New InvalidOperationException("No active ChatGPT login flow to cancel.")
            End If

            Dim loginId = _viewModel.SessionState.CurrentLoginId
            Await _accountService.CancelLoginAsync(loginId, CancellationToken.None)
            NotifyLoginCanceledUi(loginId)
            SetSessionCurrentLoginId(String.Empty)
            RefreshControlStates()
            Await RefreshAccountAsync(False)
        End Function

        Private Async Function LogoutAsync() As Task
            BeginLogoutUiTransition()
            Try
                Await _accountService.LogoutAsync(CancellationToken.None)
                NotifyLogoutCompletedUi()
                SetSessionCurrentLoginId(String.Empty)
                RefreshControlStates()
                Await RefreshAccountAsync(False)
                Await Dispatcher.Yield(DispatcherPriority.Render)
                ApplyAuthenticationRequiredState()
            Finally
                EndLogoutUiTransition()
            End Try
        End Function

        Private Async Function LoginExternalTokensAsync() As Task
            Dim idToken = _viewModel.SettingsPanel.ExternalIdToken.Trim()
            Dim accessToken = _viewModel.SettingsPanel.ExternalAccessToken.Trim()

            If String.IsNullOrWhiteSpace(idToken) Then
                Throw New InvalidOperationException("Enter an external ChatGPT ID token.")
            End If

            If String.IsNullOrWhiteSpace(accessToken) Then
                Throw New InvalidOperationException("Enter an external ChatGPT access token.")
            End If

            Await _accountService.StartExternalTokenLoginAsync(idToken,
                                                               accessToken,
                                                               CancellationToken.None)
            NotifyExternalTokensAppliedUi()
            Await EnsureAuthenticatedWithRetryAsync(allowAutoLogin:=False)
        End Function

        Private Async Function HandleChatgptTokenRefreshAsync(request As RpcServerRequest) As Task
            Dim idToken = _viewModel.SettingsPanel.ExternalIdToken.Trim()
            Dim accessToken = _viewModel.SettingsPanel.ExternalAccessToken.Trim()

            If String.IsNullOrWhiteSpace(idToken) OrElse String.IsNullOrWhiteSpace(accessToken) Then
                Await CurrentClient().SendErrorAsync(request.Id,
                                                     -32001,
                                                     "ChatGPT auth token refresh requested, but external idToken/accessToken are not configured in the UI.")
                AppendSystemMessage("Could not refresh external ChatGPT token: missing idToken/accessToken.")
                ShowStatus("Could not refresh external token: missing idToken/accessToken.", isError:=True)
                Return
            End If

            Dim response As New JsonObject()
            response("idToken") = idToken
            response("accessToken") = accessToken

            Await CurrentClient().SendResultAsync(request.Id, response)
            AppendSystemMessage("Provided refreshed external ChatGPT token to Codex.")
            ShowStatus("Provided refreshed external ChatGPT token.")
        End Function

        Private Shared Function FormatAccountLabel(accountObject As JsonObject, requiresOpenAiAuth As Boolean) As String
            If accountObject Is Nothing Then
                If requiresOpenAiAuth Then
                    Return "Account: signed out (OpenAI auth required)"
                End If

                Return "Account: signed out"
            End If

            Dim accountType = GetPropertyString(accountObject, "type")
            Select Case accountType
                Case "apiKey"
                    Return "Account: API key"
                Case "chatgpt"
                    Dim email = GetPropertyString(accountObject, "email")
                    Dim planType = GetPropertyString(accountObject, "planType")
                    Return $"Account: ChatGPT {email} ({planType})"
                Case Else
                    Return $"Account: {accountType}"
            End Select
        End Function

        Private Sub InitializeReliabilityLayer()
            _watchdogTimer.Interval = TimeSpan.FromSeconds(5)
            AddHandler _watchdogTimer.Tick, AddressOf OnWatchdogTimerTick
            _watchdogTimer.Start()

            _reconnectUiTimer.Interval = TimeSpan.FromSeconds(1)
            AddHandler _reconnectUiTimer.Tick, AddressOf OnReconnectUiTimerTick
            _reconnectUiTimer.Start()
            UpdateReconnectCountdownUi()
        End Sub

        Private Sub MarkRpcActivity()
            SetSessionLastActivity(DateTimeOffset.UtcNow)
        End Sub

        Private Sub CancelReconnect()
            Dim cts = _reconnectCts
            _reconnectCts = Nothing

            If cts IsNot Nothing Then
                Try
                    cts.Cancel()
                Catch
                End Try

                cts.Dispose()
            End If

            ApplySessionReconnectIdleState()
            UpdateReconnectCountdownUi()
        End Sub

        Private Async Function ReconnectNowAsync() As Task
            TryBeginManualReconnect("manual reconnect requested")
            Await Task.CompletedTask
        End Function

        Private Sub BeginReconnect(reason As String, Optional force As Boolean = False)
            If Not TryPrepareReconnectStart(force) Then
                Return
            End If

            StartReconnectLoop(reason)
        End Sub

        Private Sub NotifyReconnectAttemptStarted(reason As String,
                                                 currentAttempt As Integer,
                                                 totalAttempts As Integer)
            ShowStatus($"Reconnect attempt {currentAttempt}/{totalAttempts}: {reason}", isError:=True)
            UpdateReconnectCountdownUi()
        End Sub

        Private Sub NotifyReconnectAttemptFailed(currentAttempt As Integer, ex As Exception)
            AppendSystemMessage($"Reconnect attempt {currentAttempt} failed: {ex.Message}")
            ShowStatus($"Reconnect attempt {currentAttempt} failed.", isError:=True)
        End Sub

        Private Sub NotifyReconnectSucceeded()
            ShowStatus("Reconnected successfully.", displayToast:=True)
            AppendSystemMessage("Reconnected successfully.")
        End Sub

        Private Sub HandleReconnectTerminalFailureUi()
            ApplySessionReconnectTerminalFailureState()
            UpdateReconnectCountdownUi()
            ShowStatus("Auto-reconnect failed. Please reconnect manually.",
                       isError:=True,
                       displayToast:=True)
        End Sub

        Private Sub HandleReconnectCanceledUi()
            ApplySessionReconnectCanceledState()
            UpdateReconnectCountdownUi()
            ShowStatus("Auto-reconnect canceled.")
        End Sub

        Private Sub FinalizeReconnectLoopUi(reconnectCts As CancellationTokenSource)
            If ReferenceEquals(_reconnectCts, reconnectCts) Then
                _reconnectCts = Nothing
            End If

            reconnectCts.Dispose()
            ApplySessionReconnectIdleState()
            UpdateReconnectCountdownUi()
        End Sub

        Private Sub OnWatchdogTimerTick(sender As Object, e As EventArgs)
            SyncSessionStateViewModel()
            Dim session = _viewModel.SessionState

            If Not session.ConnectionExpected Then
                Return
            End If

            If Not session.IsConnected Then
                BeginReconnect("watchdog detected a disconnected client")
                Return
            End If

            Dim inactiveFor = DateTimeOffset.UtcNow - session.LastActivityUtc
            If inactiveFor < TimeSpan.FromSeconds(90) Then
                Return
            End If

            If DateTimeOffset.UtcNow - _lastWatchdogWarningUtc < TimeSpan.FromSeconds(30) Then
                Return
            End If

            _lastWatchdogWarningUtc = DateTimeOffset.UtcNow
            ShowStatus($"No app-server activity for {CInt(inactiveFor.TotalSeconds)}s. Monitoring connection.",
                       isError:=True)
        End Sub

        Private Sub OnReconnectUiTimerTick(sender As Object, e As EventArgs)
            UpdateReconnectCountdownUi()
        End Sub

        Private Sub UpdateReconnectCountdownUi()
            SyncSessionStateViewModel()
            Dim session = _viewModel.SessionState
            _viewModel.SettingsPanel.ReconnectCountdownText =
                session.BuildReconnectCountdownText(_viewModel.SettingsPanel.AutoReconnect, DateTimeOffset.UtcNow)
        End Sub
    End Class
End Namespace
