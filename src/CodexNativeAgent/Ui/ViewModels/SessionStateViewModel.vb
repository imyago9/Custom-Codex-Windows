Imports CodexNativeAgent.Ui.Mvvm

Namespace CodexNativeAgent.Ui.ViewModels
    Public NotInheritable Class SessionStateViewModel
        Inherits ViewModelBase

        Private _isConnected As Boolean
        Private _isAuthenticated As Boolean
        Private _connectionExpected As Boolean
        Private _isReconnectInProgress As Boolean
        Private _reconnectAttempt As Integer
        Private _nextReconnectAttemptUtc As DateTimeOffset?
        Private _lastActivityUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Private _currentLoginId As String = String.Empty
        Private _currentThreadId As String = String.Empty
        Private _currentTurnId As String = String.Empty
        Private _processId As Integer

        Public Property IsConnected As Boolean
            Get
                Return _isConnected
            End Get
            Set(value As Boolean)
                If SetProperty(_isConnected, value) Then
                    RaisePropertyChanged(NameOf(IsConnectedAndAuthenticated))
                    RaisePropertyChanged(NameOf(ConnectionStateLabel))
                    RaisePropertyChanged(NameOf(CanConnect))
                    RaisePropertyChanged(NameOf(CanDisconnect))
                    RaisePropertyChanged(NameOf(CanReconnectNow))
                    RaisePropertyChanged(NameOf(CanAccountRead))
                    RaisePropertyChanged(NameOf(CanLoginApiKey))
                    RaisePropertyChanged(NameOf(CanLoginChatGpt))
                    RaisePropertyChanged(NameOf(CanLoginExternalTokens))
                    RaisePropertyChanged(NameOf(CanLogout))
                    RaisePropertyChanged(NameOf(CanReadRateLimits))
                    RaisePropertyChanged(NameOf(CanCancelLogin))
                End If
            End Set
        End Property

        Public Property IsAuthenticated As Boolean
            Get
                Return _isAuthenticated
            End Get
            Set(value As Boolean)
                If SetProperty(_isAuthenticated, value) Then
                    RaisePropertyChanged(NameOf(IsConnectedAndAuthenticated))
                    RaisePropertyChanged(NameOf(CanLogout))
                    RaisePropertyChanged(NameOf(CanReadRateLimits))
                End If
            End Set
        End Property

        Public ReadOnly Property IsConnectedAndAuthenticated As Boolean
            Get
                Return _isConnected AndAlso _isAuthenticated
            End Get
        End Property

        Public ReadOnly Property ConnectionStateLabel As String
            Get
                Return If(_isConnected, "Connected", "Disconnected")
            End Get
        End Property

        Public ReadOnly Property CanConnect As Boolean
            Get
                Return Not _isConnected
            End Get
        End Property

        Public ReadOnly Property CanDisconnect As Boolean
            Get
                Return _isConnected
            End Get
        End Property

        Public ReadOnly Property CanReconnectNow As Boolean
            Get
                Return Not _isConnected
            End Get
        End Property

        Public ReadOnly Property CanAccountRead As Boolean
            Get
                Return _isConnected
            End Get
        End Property

        Public ReadOnly Property CanLoginApiKey As Boolean
            Get
                Return _isConnected
            End Get
        End Property

        Public ReadOnly Property CanLoginChatGpt As Boolean
            Get
                Return _isConnected
            End Get
        End Property

        Public ReadOnly Property CanLoginExternalTokens As Boolean
            Get
                Return _isConnected
            End Get
        End Property

        Public ReadOnly Property CanLogout As Boolean
            Get
                Return IsConnectedAndAuthenticated
            End Get
        End Property

        Public ReadOnly Property CanReadRateLimits As Boolean
            Get
                Return IsConnectedAndAuthenticated
            End Get
        End Property

        Public Property ConnectionExpected As Boolean
            Get
                Return _connectionExpected
            End Get
            Set(value As Boolean)
                SetProperty(_connectionExpected, value)
            End Set
        End Property

        Public Property IsReconnectInProgress As Boolean
            Get
                Return _isReconnectInProgress
            End Get
            Set(value As Boolean)
                SetProperty(_isReconnectInProgress, value)
            End Set
        End Property

        Public Property ReconnectAttempt As Integer
            Get
                Return _reconnectAttempt
            End Get
            Set(value As Integer)
                SetProperty(_reconnectAttempt, Math.Max(0, value))
            End Set
        End Property

        Public Property NextReconnectAttemptUtc As DateTimeOffset?
            Get
                Return _nextReconnectAttemptUtc
            End Get
            Set(value As DateTimeOffset?)
                SetProperty(_nextReconnectAttemptUtc, value)
            End Set
        End Property

        Public Property LastActivityUtc As DateTimeOffset
            Get
                Return _lastActivityUtc
            End Get
            Set(value As DateTimeOffset)
                SetProperty(_lastActivityUtc, value)
            End Set
        End Property

        Public Property CurrentLoginId As String
            Get
                Return _currentLoginId
            End Get
            Set(value As String)
                If SetProperty(_currentLoginId, If(value, String.Empty)) Then
                    RaisePropertyChanged(NameOf(HasActiveLogin))
                    RaisePropertyChanged(NameOf(CanCancelLogin))
                End If
            End Set
        End Property

        Public ReadOnly Property HasActiveLogin As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(_currentLoginId)
            End Get
        End Property

        Public ReadOnly Property CanCancelLogin As Boolean
            Get
                Return _isConnected AndAlso HasActiveLogin
            End Get
        End Property

        Public Property CurrentThreadId As String
            Get
                Return _currentThreadId
            End Get
            Set(value As String)
                If SetProperty(_currentThreadId, If(value, String.Empty)) Then
                    RaisePropertyChanged(NameOf(HasCurrentThread))
                End If
            End Set
        End Property

        Public ReadOnly Property HasCurrentThread As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(_currentThreadId)
            End Get
        End Property

        Public Property CurrentTurnId As String
            Get
                Return _currentTurnId
            End Get
            Set(value As String)
                If SetProperty(_currentTurnId, If(value, String.Empty)) Then
                    RaisePropertyChanged(NameOf(HasCurrentTurn))
                End If
            End Set
        End Property

        Public ReadOnly Property HasCurrentTurn As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(_currentTurnId)
            End Get
        End Property

        Public Property ProcessId As Integer
            Get
                Return _processId
            End Get
            Set(value As Integer)
                SetProperty(_processId, Math.Max(0, value))
            End Set
        End Property

        Public Function BuildReconnectCountdownText(autoReconnectEnabled As Boolean,
                                                    nowUtc As DateTimeOffset) As String
            If _isConnected Then
                Return "Reconnect: connected."
            End If

            If _isReconnectInProgress Then
                If _nextReconnectAttemptUtc.HasValue Then
                    Dim remaining = _nextReconnectAttemptUtc.Value - nowUtc
                    Dim secondsLeft = Math.Max(0, CInt(Math.Ceiling(remaining.TotalSeconds)))
                    Return $"Reconnect: next attempt in {secondsLeft}s."
                End If

                Dim attemptText = If(_reconnectAttempt > 0, _reconnectAttempt.ToString(), "?")
                Return $"Reconnect: attempt {attemptText} running..."
            End If

            If _connectionExpected AndAlso autoReconnectEnabled Then
                Return "Reconnect: standing by."
            End If

            Return "Reconnect: not scheduled."
        End Function

        Public Sub ApplyConnectionEstablished(nowUtc As DateTimeOffset)
            ConnectionExpected = True
            LastActivityUtc = nowUtc
            NextReconnectAttemptUtc = Nothing
        End Sub

        Public Sub ApplyReconnectLoopStarted()
            ConnectionExpected = True
            IsReconnectInProgress = True
            ReconnectAttempt = 0
            NextReconnectAttemptUtc = Nothing
        End Sub

        Public Sub ApplyReconnectAttemptStarted(attempt As Integer)
            ReconnectAttempt = Math.Max(0, attempt)
            NextReconnectAttemptUtc = Nothing
        End Sub

        Public Sub ApplyReconnectRetryScheduled(nextAttemptUtc As DateTimeOffset)
            NextReconnectAttemptUtc = nextAttemptUtc
        End Sub

        Public Sub ApplyReconnectTerminalFailureState()
            ConnectionExpected = False
            NextReconnectAttemptUtc = Nothing
        End Sub

        Public Sub ApplyReconnectCanceledState()
            NextReconnectAttemptUtc = Nothing
        End Sub

        Public Sub ApplyReconnectIdleState()
            IsReconnectInProgress = False
            ReconnectAttempt = 0
            NextReconnectAttemptUtc = Nothing
        End Sub
    End Class
End Namespace
