Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports CodexNativeAgent.AppServer
Imports CodexNativeAgent.Services
Imports CodexNativeAgent.Ui.Coordinators

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
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
            session.CurrentThreadId = _currentThreadId
            session.CurrentTurnId = _currentTurnId
            session.ProcessId = If(_client Is Nothing, 0, _client.ProcessId)
        End Sub

        Private Sub SyncCurrentTurnFromRuntimeStore(Optional keepExistingWhenRuntimeIsIdle As Boolean = False)
            If _sessionNotificationCoordinator Is Nothing Then
                Return
            End If

            Dim normalizedThreadId = If(_currentThreadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                If Not keepExistingWhenRuntimeIsIdle Then
                    _currentTurnId = String.Empty
                End If
                Return
            End If

            Dim runtimeStore = _sessionNotificationCoordinator.RuntimeStore
            Dim activeTurnId = runtimeStore.GetActiveTurnId(normalizedThreadId, _currentTurnId)
            If Not String.IsNullOrWhiteSpace(activeTurnId) Then
                _currentTurnId = activeTurnId
                Return
            End If

            If keepExistingWhenRuntimeIsIdle Then
                Dim latestTurnId = runtimeStore.GetLatestTurnId(normalizedThreadId)
                If String.IsNullOrWhiteSpace(latestTurnId) Then
                    Return
                End If
            End If

            _currentTurnId = String.Empty
        End Sub

        Private Function HasActiveRuntimeTurnForCurrentThread() As Boolean
            If _sessionNotificationCoordinator Is Nothing Then
                Return False
            End If

            Dim normalizedThreadId = If(_currentThreadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return False
            End If

            Return _sessionNotificationCoordinator.RuntimeStore.HasActiveTurn(normalizedThreadId)
        End Function

        Private Function RuntimeHasTurnHistoryForCurrentThread() As Boolean
            If _sessionNotificationCoordinator Is Nothing Then
                Return False
            End If

            Dim normalizedThreadId = If(_currentThreadId, String.Empty).Trim()
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
            _currentThreadId = String.Empty
            _currentTurnId = String.Empty
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
            _threadLiveSessionRegistry.Clear()
            _sessionNotificationCoordinator.ResetStreamingAgentItems()
            _pendingLocalUserEchoes.Clear()
            _turnWorkflowCoordinator.ResetApprovalState()
            _viewModel.ThreadsPanel.Items.Clear()
            _viewModel.ThreadsPanel.SelectedListItem = Nothing
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
        End Sub

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
            SyncThreadContentLoadingFromThreadsPanel()
            Return loadVersion
        End Function

        Private Function IsCurrentThreadSelectionLoadUiState(loadVersion As Integer,
                                                             threadId As String) As Boolean
            Return _viewModel.ThreadsPanel.IsCurrentThreadSelectionLoad(loadVersion, threadId)
        End Function

        Private Function TryCompleteThreadSelectionLoadUiState(loadVersion As Integer) As Boolean
            Dim completed = _viewModel.ThreadsPanel.TryCompleteThreadSelectionLoad(loadVersion)
            If completed Then
                SyncThreadContentLoadingFromThreadsPanel()
            End If

            Return completed
        End Function

        Private Sub ResetThreadSelectionLoadUiState(Optional hideTranscriptLoader As Boolean = False)
            _viewModel.ThreadsPanel.CancelThreadSelectionLoadState()
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
            BeginUserDisconnectSessionTransition()

            Dim client = DetachCurrentClient()
            If client IsNot Nothing Then
                Await DisconnectClientInternalAsync(client,
                                                    "Disconnected by user.",
                                                    CancellationToken.None)
            End If

            ResetDisconnectedUiState("Disconnected.", isError:=False, displayToast:=True)
        End Function

        Private Async Function DisconnectClientInternalAsync(client As CodexAppServerClient,
                                                             reason As String,
                                                             cancellationToken As CancellationToken) As Task
            If client Is Nothing Then
                Return
            End If

            _disconnecting = True
            Try
                Await _connectionService.StopAsync(client, reason, cancellationToken)
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
            Dim visibleThreadIdBeforeDispatch = If(_currentThreadId, String.Empty).Trim()
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
            Dim visibleThreadIdBeforeDispatch = If(_currentThreadId, String.Empty).Trim()
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
                                                      If(_currentThreadId, String.Empty).Trim(),
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
                    Continue For
                End If

                If Not HandleRuntimeEventForThreadVisibility(resolvedThreadId,
                                                             resolvedTurnId,
                                                             visibleThreadIdForRuntimeRouting,
                                                             dispatch.MethodName) Then
                    Continue For
                End If

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
                ScrollTranscriptToBottom()
            End If
        End Sub

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

            Dim visibleThreadIdForRuntimeRouting = If(_currentThreadId, String.Empty).Trim()
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
                ScrollTranscriptToBottom()
            End If

            RefreshThreadRuntimeIndicatorsIfNeeded()
        End Sub

        Private Sub ApplyApprovalResolutionDispatchResult(dispatch As SessionNotificationCoordinator.ApprovalResolutionDispatchResult)
            If dispatch Is Nothing Then
                Return
            End If

            ApplyProtocolDispatchMessages(dispatch.ProtocolMessages)

            Dim visibleThreadIdForRuntimeRouting = If(_currentThreadId, String.Empty).Trim()
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
                ScrollTranscriptToBottom()
            End If

            RefreshThreadRuntimeIndicatorsIfNeeded()
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
                _threadLiveSessionRegistry.MarkBound(normalizedThreadId, _currentTurnId)
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
                baselineVisibleThreadId = If(_currentThreadId, String.Empty).Trim()
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
                If ShouldInitializeWorkspaceAfterAuthentication(wasAuthenticated) Then
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
                ShowStatus("Connected and authenticated.")
            Finally
                EndWorkspaceBootstrapAfterAuthentication()
            End Try
        End Function

        Private Sub ApplyAuthenticationRequiredState(Optional showPrompt As Boolean = False)
            ResetWorkspaceForAuthenticationRequired()
            FinalizeAuthenticationRequiredWorkspaceUi()
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
            Await _accountService.LogoutAsync(CancellationToken.None)
            NotifyLogoutCompletedUi()
            SetSessionCurrentLoginId(String.Empty)
            RefreshControlStates()
            Await RefreshAccountAsync(False)
            ApplyAuthenticationRequiredState()
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
