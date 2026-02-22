Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Input
Imports System.Text.Json.Nodes
Imports CodexNativeAgent.Services
Imports CodexNativeAgent.Ui.Mvvm
Imports CodexNativeAgent.Ui.ViewModels

Namespace CodexNativeAgent.Ui.Coordinators
    Public NotInheritable Class SessionCoordinator
        Public NotInheritable Class ReconnectAttemptStartedEventArgs
            Inherits EventArgs

            Public Sub New(reason As String, currentAttempt As Integer, totalAttempts As Integer)
                Me.Reason = If(reason, String.Empty)
                Me.CurrentAttempt = Math.Max(0, currentAttempt)
                Me.TotalAttempts = Math.Max(0, totalAttempts)
            End Sub

            Public ReadOnly Property Reason As String
            Public ReadOnly Property CurrentAttempt As Integer
            Public ReadOnly Property TotalAttempts As Integer
        End Class

        Public NotInheritable Class ReconnectAttemptFailedEventArgs
            Inherits EventArgs

            Public Sub New(currentAttempt As Integer, [error] As Exception)
                Me.CurrentAttempt = Math.Max(0, currentAttempt)
                Me.Error = [error]
            End Sub

            Public ReadOnly Property CurrentAttempt As Integer
            Public ReadOnly Property [Error] As Exception
        End Class

        Public NotInheritable Class ReconnectRetryScheduledEventArgs
            Inherits EventArgs

            Public Sub New(nextAttemptUtc As DateTimeOffset)
                Me.NextAttemptUtc = nextAttemptUtc
            End Sub

            Public ReadOnly Property NextAttemptUtc As DateTimeOffset
        End Class

        Public NotInheritable Class ReconnectFinalizeEventArgs
            Inherits EventArgs

            Public Sub New(reconnectCancellationSource As CancellationTokenSource)
                Me.ReconnectCancellationSource = reconnectCancellationSource
            End Sub

            Public ReadOnly Property ReconnectCancellationSource As CancellationTokenSource
        End Class

        Private ReadOnly _viewModel As MainWindowViewModel
        Private ReadOnly _runUiActionAsync As Func(Of Func(Of Task), Task)
        Private ReadOnly _connectAsync As Func(Of Task)
        Private ReadOnly _disconnectAsync As Func(Of Task)
        Private ReadOnly _reconnectNowAsync As Func(Of Task)
        Private ReadOnly _refreshAuthenticationGateAsync As Func(Of Task)
        Private ReadOnly _loginApiKeyAsync As Func(Of Task)
        Private ReadOnly _loginChatGptAsync As Func(Of Task)
        Private ReadOnly _cancelLoginAsync As Func(Of Task)
        Private ReadOnly _logoutAsync As Func(Of Task)
        Private ReadOnly _readRateLimitsAsync As Func(Of Task)
        Private ReadOnly _loginExternalTokensAsync As Func(Of Task)

        Public Event ReconnectAttemptStarted As EventHandler(Of ReconnectAttemptStartedEventArgs)
        Public Event ReconnectAttemptFailed As EventHandler(Of ReconnectAttemptFailedEventArgs)
        Public Event ReconnectSucceeded As EventHandler
        Public Event ReconnectRetryScheduled As EventHandler(Of ReconnectRetryScheduledEventArgs)
        Public Event ReconnectTerminalFailure As EventHandler
        Public Event ReconnectCanceled As EventHandler
        Public Event ReconnectFinalizing As EventHandler(Of ReconnectFinalizeEventArgs)

        Public Sub New(viewModel As MainWindowViewModel,
                       runUiActionAsync As Func(Of Func(Of Task), Task),
                       connectAsync As Func(Of Task),
                       disconnectAsync As Func(Of Task),
                       reconnectNowAsync As Func(Of Task),
                       refreshAuthenticationGateAsync As Func(Of Task),
                       loginApiKeyAsync As Func(Of Task),
                       loginChatGptAsync As Func(Of Task),
                       cancelLoginAsync As Func(Of Task),
                       logoutAsync As Func(Of Task),
                       readRateLimitsAsync As Func(Of Task),
                       loginExternalTokensAsync As Func(Of Task))
            If viewModel Is Nothing Then Throw New ArgumentNullException(NameOf(viewModel))
            If runUiActionAsync Is Nothing Then Throw New ArgumentNullException(NameOf(runUiActionAsync))
            If connectAsync Is Nothing Then Throw New ArgumentNullException(NameOf(connectAsync))
            If disconnectAsync Is Nothing Then Throw New ArgumentNullException(NameOf(disconnectAsync))
            If reconnectNowAsync Is Nothing Then Throw New ArgumentNullException(NameOf(reconnectNowAsync))
            If refreshAuthenticationGateAsync Is Nothing Then Throw New ArgumentNullException(NameOf(refreshAuthenticationGateAsync))
            If loginApiKeyAsync Is Nothing Then Throw New ArgumentNullException(NameOf(loginApiKeyAsync))
            If loginChatGptAsync Is Nothing Then Throw New ArgumentNullException(NameOf(loginChatGptAsync))
            If cancelLoginAsync Is Nothing Then Throw New ArgumentNullException(NameOf(cancelLoginAsync))
            If logoutAsync Is Nothing Then Throw New ArgumentNullException(NameOf(logoutAsync))
            If readRateLimitsAsync Is Nothing Then Throw New ArgumentNullException(NameOf(readRateLimitsAsync))
            If loginExternalTokensAsync Is Nothing Then Throw New ArgumentNullException(NameOf(loginExternalTokensAsync))

            _viewModel = viewModel
            _runUiActionAsync = runUiActionAsync
            _connectAsync = connectAsync
            _disconnectAsync = disconnectAsync
            _reconnectNowAsync = reconnectNowAsync
            _refreshAuthenticationGateAsync = refreshAuthenticationGateAsync
            _loginApiKeyAsync = loginApiKeyAsync
            _loginChatGptAsync = loginChatGptAsync
            _cancelLoginAsync = cancelLoginAsync
            _logoutAsync = logoutAsync
            _readRateLimitsAsync = readRateLimitsAsync
            _loginExternalTokensAsync = loginExternalTokensAsync
        End Sub

        Public Sub BindSettingsCommands()
            Dim settings = _viewModel.SettingsPanel
            Dim session = _viewModel.SessionState

            settings.ConnectCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(_connectAsync)
                End Function,
                Function() session.CanConnect)

            settings.DisconnectCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(_disconnectAsync)
                End Function,
                Function() session.CanDisconnect)

            settings.ReconnectNowCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(_reconnectNowAsync)
                End Function,
                Function() session.CanReconnectNow)

            settings.AccountReadCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(_refreshAuthenticationGateAsync)
                End Function,
                Function() session.CanAccountRead)

            settings.LoginApiKeyCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(_loginApiKeyAsync)
                End Function,
                Function() session.CanLoginApiKey)

            settings.LoginChatGptCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(_loginChatGptAsync)
                End Function,
                Function() session.CanLoginChatGpt)

            settings.CancelLoginCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(_cancelLoginAsync)
                End Function,
                Function() session.CanCancelLogin)

            settings.LogoutCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(_logoutAsync)
                End Function,
                Function() session.CanLogout)

            settings.ReadRateLimitsCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(_readRateLimitsAsync)
                End Function,
                Function() session.CanReadRateLimits)

            settings.LoginExternalTokensCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(_loginExternalTokensAsync)
                End Function,
                Function() session.CanLoginExternalTokens)
        End Sub

        Public Sub RefreshCommandCanExecuteStates()
            Dim settings = _viewModel.SettingsPanel
            RaiseAsyncCanExecuteChanged(settings.ConnectCommand)
            RaiseAsyncCanExecuteChanged(settings.DisconnectCommand)
            RaiseAsyncCanExecuteChanged(settings.ReconnectNowCommand)
            RaiseAsyncCanExecuteChanged(settings.AccountReadCommand)
            RaiseAsyncCanExecuteChanged(settings.LoginApiKeyCommand)
            RaiseAsyncCanExecuteChanged(settings.LoginChatGptCommand)
            RaiseAsyncCanExecuteChanged(settings.CancelLoginCommand)
            RaiseAsyncCanExecuteChanged(settings.LogoutCommand)
            RaiseAsyncCanExecuteChanged(settings.ReadRateLimitsCommand)
            RaiseAsyncCanExecuteChanged(settings.LoginExternalTokensCommand)
        End Sub

        Public Async Function RunReconnectLoopAsync(reason As String,
                                                    reconnectCts As CancellationTokenSource,
                                                    connectAttemptAsync As Func(Of CancellationToken, Task)) As Task
            If reconnectCts Is Nothing Then
                Throw New ArgumentNullException(NameOf(reconnectCts))
            End If

            If connectAttemptAsync Is Nothing Then
                Throw New ArgumentNullException(NameOf(connectAttemptAsync))
            End If

            Dim token = reconnectCts.Token
            Dim delays As Integer() = {2, 5, 10, 20, 30}

            Try
                For attempt = 1 To delays.Length + 1
                    token.ThrowIfCancellationRequested()

                    Dim currentAttempt = attempt
                    RaiseEvent ReconnectAttemptStarted(Me, New ReconnectAttemptStartedEventArgs(reason, currentAttempt, delays.Length + 1))

                    Dim delayBeforeRetry As TimeSpan? = Nothing
                    Try
                        Await connectAttemptAsync(token).ConfigureAwait(False)
                        RaiseEvent ReconnectSucceeded(Me, EventArgs.Empty)
                        Return
                    Catch ex As OperationCanceledException
                        Throw
                    Catch ex As Exception
                        RaiseEvent ReconnectAttemptFailed(Me, New ReconnectAttemptFailedEventArgs(currentAttempt, ex))

                        If currentAttempt <= delays.Length Then
                            delayBeforeRetry = TimeSpan.FromSeconds(delays(currentAttempt - 1))
                        End If
                    End Try

                    If delayBeforeRetry.HasValue Then
                        Dim nextAttemptUtc = DateTimeOffset.UtcNow.Add(delayBeforeRetry.Value)
                        RaiseEvent ReconnectRetryScheduled(Me, New ReconnectRetryScheduledEventArgs(nextAttemptUtc))
                        Await Task.Delay(delayBeforeRetry.Value, token).ConfigureAwait(False)
                    End If
                Next

                RaiseEvent ReconnectTerminalFailure(Me, EventArgs.Empty)
            Catch ex As OperationCanceledException
                RaiseEvent ReconnectCanceled(Me, EventArgs.Empty)
            Finally
                RaiseEvent ReconnectFinalizing(Me, New ReconnectFinalizeEventArgs(reconnectCts))
            End Try
        End Function

        Public Async Function RunAuthenticationRetryAsync(allowAutoLogin As Boolean,
                                                          ensureAttemptAsync As Func(Of Boolean, Boolean, Task(Of Boolean))) As Task(Of Boolean)
            If ensureAttemptAsync Is Nothing Then
                Throw New ArgumentNullException(NameOf(ensureAttemptAsync))
            End If

            Dim delaysMs As Integer() = {0, 500, 1200, 2500}

            For attempt = 0 To delaysMs.Length - 1
                If delaysMs(attempt) > 0 Then
                    Await Task.Delay(delaysMs(attempt)).ConfigureAwait(True)
                End If

                Dim useAutoLogin = allowAutoLogin AndAlso attempt = 0
                Dim showAuthPrompt = attempt = delaysMs.Length - 1
                Dim hasAccount = Await ensureAttemptAsync(useAutoLogin, showAuthPrompt).ConfigureAwait(True)
                If hasAccount Then
                    Return True
                End If
            Next

            Return False
        End Function

        Public Async Function ReadAccountAndApplyAsync(refreshToken As Boolean,
                                                       readAccountAsync As Func(Of Boolean, CancellationToken, Task(Of AccountReadResult)),
                                                       applyAccountResultUi As Func(Of AccountReadResult, Boolean)) As Task(Of Boolean)
            If readAccountAsync Is Nothing Then
                Throw New ArgumentNullException(NameOf(readAccountAsync))
            End If

            If applyAccountResultUi Is Nothing Then
                Throw New ArgumentNullException(NameOf(applyAccountResultUi))
            End If

            Dim result = Await readAccountAsync(refreshToken, CancellationToken.None).ConfigureAwait(True)
            Return applyAccountResultUi(result)
        End Function

        Public Async Function ReadRateLimitsAndApplyAsync(readRateLimitsAsync As Func(Of CancellationToken, Task(Of JsonNode)),
                                                          applyRateLimitsUi As Action(Of JsonNode)) As Task
            If readRateLimitsAsync Is Nothing Then
                Throw New ArgumentNullException(NameOf(readRateLimitsAsync))
            End If

            If applyRateLimitsUi Is Nothing Then
                Throw New ArgumentNullException(NameOf(applyRateLimitsUi))
            End If

            Dim response = Await readRateLimitsAsync(CancellationToken.None).ConfigureAwait(True)
            applyRateLimitsUi(response)
        End Function

        Private Shared Sub RaiseAsyncCanExecuteChanged(command As ICommand)
            Dim asyncCommand = TryCast(command, AsyncRelayCommand)
            If asyncCommand Is Nothing Then
                Return
            End If

            asyncCommand.RaiseCanExecuteChanged()
        End Sub
    End Class
End Namespace
