Imports System.Windows.Input
Imports CodexNativeAgent.Ui.Mvvm

Namespace CodexNativeAgent.Ui.ViewModels
    Public NotInheritable Class SettingsPanelViewModel
        Inherits ViewModelBase

        Private _themeStateText As String = "Current: Light"
        Private _themeToggleButtonText As String = "Switch to Dark Mode"
        Private _densityIndex As Integer
        Private _transcriptScaleIndex As Integer = 2

        Private _codexPath As String = String.Empty
        Private _serverArgs As String = String.Empty
        Private _workingDir As String = String.Empty
        Private _windowsCodexHome As String = String.Empty
        Private _apiKey As String = String.Empty
        Private _externalAccessToken As String = String.Empty
        Private _externalIdToken As String = String.Empty
        Private _externalAccountId As String = String.Empty
        Private _externalPlanType As String = String.Empty
        Private _rememberApiKey As Boolean
        Private _autoLoginApiKey As Boolean
        Private _autoReconnect As Boolean
        Private _disableWorkspaceHintOverlay As Boolean
        Private _disableConnectionInitializedToast As Boolean
        Private _disableThreadsPanelHints As Boolean
        Private _showEventDotsInTranscript As Boolean
        Private _showSystemDotsInTranscript As Boolean
        Private _areConnectionFieldsEditable As Boolean = True

        Private _reconnectCountdownText As String = "Reconnect: not scheduled."
        Private _accountStateText As String = "Account: unknown"
        Private _rateLimitsText As String = "No rate-limit data loaded yet."
        Private _canExportDiagnostics As Boolean = True

        Private _toggleThemeCommand As ICommand
        Private _connectCommand As ICommand
        Private _disconnectCommand As ICommand
        Private _reconnectNowCommand As ICommand
        Private _exportDiagnosticsCommand As ICommand
        Private _accountReadCommand As ICommand
        Private _loginApiKeyCommand As ICommand
        Private _loginChatGptCommand As ICommand
        Private _cancelLoginCommand As ICommand
        Private _logoutCommand As ICommand
        Private _readRateLimitsCommand As ICommand
        Private _loginExternalTokensCommand As ICommand

        Public Property ThemeStateText As String
            Get
                Return _themeStateText
            End Get
            Set(value As String)
                SetProperty(_themeStateText, If(value, String.Empty))
            End Set
        End Property

        Public Property ThemeToggleButtonText As String
            Get
                Return _themeToggleButtonText
            End Get
            Set(value As String)
                SetProperty(_themeToggleButtonText, If(value, String.Empty))
            End Set
        End Property

        Public Property DensityIndex As Integer
            Get
                Return _densityIndex
            End Get
            Set(value As Integer)
                SetProperty(_densityIndex, Math.Max(0, value))
            End Set
        End Property

        Public Property TranscriptScaleIndex As Integer
            Get
                Return _transcriptScaleIndex
            End Get
            Set(value As Integer)
                SetProperty(_transcriptScaleIndex, Math.Max(0, value))
            End Set
        End Property

        Public Property CodexPath As String
            Get
                Return _codexPath
            End Get
            Set(value As String)
                SetProperty(_codexPath, If(value, String.Empty))
            End Set
        End Property

        Public Property ServerArgs As String
            Get
                Return _serverArgs
            End Get
            Set(value As String)
                SetProperty(_serverArgs, If(value, String.Empty))
            End Set
        End Property

        Public Property WorkingDir As String
            Get
                Return _workingDir
            End Get
            Set(value As String)
                SetProperty(_workingDir, If(value, String.Empty))
            End Set
        End Property

        Public Property WindowsCodexHome As String
            Get
                Return _windowsCodexHome
            End Get
            Set(value As String)
                SetProperty(_windowsCodexHome, If(value, String.Empty))
            End Set
        End Property

        Public Property ApiKey As String
            Get
                Return _apiKey
            End Get
            Set(value As String)
                SetProperty(_apiKey, If(value, String.Empty))
            End Set
        End Property

        Public Property ExternalAccessToken As String
            Get
                Return _externalAccessToken
            End Get
            Set(value As String)
                SetProperty(_externalAccessToken, If(value, String.Empty))
            End Set
        End Property

        Public Property ExternalIdToken As String
            Get
                Return _externalIdToken
            End Get
            Set(value As String)
                SetProperty(_externalIdToken, If(value, String.Empty))
            End Set
        End Property

        Public Property ExternalAccountId As String
            Get
                Return _externalAccountId
            End Get
            Set(value As String)
                SetProperty(_externalAccountId, If(value, String.Empty))
            End Set
        End Property

        Public Property ExternalPlanType As String
            Get
                Return _externalPlanType
            End Get
            Set(value As String)
                SetProperty(_externalPlanType, If(value, String.Empty))
            End Set
        End Property

        Public Property RememberApiKey As Boolean
            Get
                Return _rememberApiKey
            End Get
            Set(value As Boolean)
                SetProperty(_rememberApiKey, value)
            End Set
        End Property

        Public Property AutoLoginApiKey As Boolean
            Get
                Return _autoLoginApiKey
            End Get
            Set(value As Boolean)
                SetProperty(_autoLoginApiKey, value)
            End Set
        End Property

        Public Property AutoReconnect As Boolean
            Get
                Return _autoReconnect
            End Get
            Set(value As Boolean)
                SetProperty(_autoReconnect, value)
            End Set
        End Property

        Public Property DisableWorkspaceHintOverlay As Boolean
            Get
                Return _disableWorkspaceHintOverlay
            End Get
            Set(value As Boolean)
                SetProperty(_disableWorkspaceHintOverlay, value)
            End Set
        End Property

        Public Property DisableConnectionInitializedToast As Boolean
            Get
                Return _disableConnectionInitializedToast
            End Get
            Set(value As Boolean)
                SetProperty(_disableConnectionInitializedToast, value)
            End Set
        End Property

        Public Property DisableThreadsPanelHints As Boolean
            Get
                Return _disableThreadsPanelHints
            End Get
            Set(value As Boolean)
                SetProperty(_disableThreadsPanelHints, value)
            End Set
        End Property

        Public Property ShowEventDotsInTranscript As Boolean
            Get
                Return _showEventDotsInTranscript
            End Get
            Set(value As Boolean)
                SetProperty(_showEventDotsInTranscript, value)
            End Set
        End Property

        Public Property ShowSystemDotsInTranscript As Boolean
            Get
                Return _showSystemDotsInTranscript
            End Get
            Set(value As Boolean)
                SetProperty(_showSystemDotsInTranscript, value)
            End Set
        End Property

        Public Property AreConnectionFieldsEditable As Boolean
            Get
                Return _areConnectionFieldsEditable
            End Get
            Set(value As Boolean)
                SetProperty(_areConnectionFieldsEditable, value)
            End Set
        End Property

        Public Property ReconnectCountdownText As String
            Get
                Return _reconnectCountdownText
            End Get
            Set(value As String)
                SetProperty(_reconnectCountdownText, If(value, String.Empty))
            End Set
        End Property

        Public Property AccountStateText As String
            Get
                Return _accountStateText
            End Get
            Set(value As String)
                SetProperty(_accountStateText, If(value, String.Empty))
            End Set
        End Property

        Public Property RateLimitsText As String
            Get
                Return _rateLimitsText
            End Get
            Set(value As String)
                SetProperty(_rateLimitsText, If(value, String.Empty))
            End Set
        End Property

        Public Property CanExportDiagnostics As Boolean
            Get
                Return _canExportDiagnostics
            End Get
            Set(value As Boolean)
                SetProperty(_canExportDiagnostics, value)
            End Set
        End Property

        Public Property ToggleThemeCommand As ICommand
            Get
                Return _toggleThemeCommand
            End Get
            Set(value As ICommand)
                SetProperty(_toggleThemeCommand, value)
            End Set
        End Property

        Public Property ConnectCommand As ICommand
            Get
                Return _connectCommand
            End Get
            Set(value As ICommand)
                SetProperty(_connectCommand, value)
            End Set
        End Property

        Public Property DisconnectCommand As ICommand
            Get
                Return _disconnectCommand
            End Get
            Set(value As ICommand)
                SetProperty(_disconnectCommand, value)
            End Set
        End Property

        Public Property ReconnectNowCommand As ICommand
            Get
                Return _reconnectNowCommand
            End Get
            Set(value As ICommand)
                SetProperty(_reconnectNowCommand, value)
            End Set
        End Property

        Public Property ExportDiagnosticsCommand As ICommand
            Get
                Return _exportDiagnosticsCommand
            End Get
            Set(value As ICommand)
                SetProperty(_exportDiagnosticsCommand, value)
            End Set
        End Property

        Public Property AccountReadCommand As ICommand
            Get
                Return _accountReadCommand
            End Get
            Set(value As ICommand)
                SetProperty(_accountReadCommand, value)
            End Set
        End Property

        Public Property LoginApiKeyCommand As ICommand
            Get
                Return _loginApiKeyCommand
            End Get
            Set(value As ICommand)
                SetProperty(_loginApiKeyCommand, value)
            End Set
        End Property

        Public Property LoginChatGptCommand As ICommand
            Get
                Return _loginChatGptCommand
            End Get
            Set(value As ICommand)
                SetProperty(_loginChatGptCommand, value)
            End Set
        End Property

        Public Property CancelLoginCommand As ICommand
            Get
                Return _cancelLoginCommand
            End Get
            Set(value As ICommand)
                SetProperty(_cancelLoginCommand, value)
            End Set
        End Property

        Public Property LogoutCommand As ICommand
            Get
                Return _logoutCommand
            End Get
            Set(value As ICommand)
                SetProperty(_logoutCommand, value)
            End Set
        End Property

        Public Property ReadRateLimitsCommand As ICommand
            Get
                Return _readRateLimitsCommand
            End Get
            Set(value As ICommand)
                SetProperty(_readRateLimitsCommand, value)
            End Set
        End Property

        Public Property LoginExternalTokensCommand As ICommand
            Get
                Return _loginExternalTokensCommand
            End Get
            Set(value As ICommand)
                SetProperty(_loginExternalTokensCommand, value)
            End Set
        End Property
    End Class
End Namespace
