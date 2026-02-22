Imports System.Windows.Input
Imports System.Windows
Imports CodexNativeAgent.Ui.Mvvm

Namespace CodexNativeAgent.Ui.ViewModels
    Public NotInheritable Class MainWindowViewModel
        Inherits ViewModelBase

        Private ReadOnly _turnComposer As New TurnComposerViewModel()
        Private ReadOnly _threadsPanel As New ThreadsPanelViewModel()
        Private ReadOnly _transcriptPanel As New TranscriptPanelViewModel()
        Private ReadOnly _approvalPanel As New ApprovalPanelViewModel()
        Private ReadOnly _settingsPanel As New SettingsPanelViewModel()
        Private ReadOnly _sessionState As New SessionStateViewModel()

        Private _statusText As String = "Ready."
        Private _currentThreadText As String = "New thread"
        Private _currentTurnText As String = "Turn: 0"
        Private _connectionStateText As String = "Disconnected"
        Private _workspaceHintText As String = "Connect to begin."
        Private _sidebarNewThreadButtonText As String = "New thread"
        Private _sidebarNewThreadToolTip As Object
        Private _sidebarMainViewVisibility As Visibility = Visibility.Visible
        Private _sidebarSettingsViewVisibility As Visibility = Visibility.Collapsed
        Private _sidebarNewThreadNavTag As String = "Active"
        Private _sidebarAutomationsNavTag As String = String.Empty
        Private _sidebarSkillsNavTag As String = String.Empty
        Private _sidebarSettingsNavTag As String = String.Empty
        Private _isSidebarNewThreadEnabled As Boolean
        Private _isSidebarAutomationsEnabled As Boolean = True
        Private _isSidebarSkillsEnabled As Boolean = True
        Private _isSidebarSettingsEnabled As Boolean = True
        Private _isSettingsBackEnabled As Boolean = True
        Private _isQuickOpenVscEnabled As Boolean = True
        Private _isQuickOpenTerminalEnabled As Boolean = True
        Private _areThreadLegacyFilterControlsEnabled As Boolean
        Private _shellSendCommand As ICommand
        Private _shellRefreshThreadsCommand As ICommand
        Private _shellRefreshModelsCommand As ICommand
        Private _shellNewThreadCommand As ICommand
        Private _shellFocusThreadSearchCommand As ICommand
        Private _shellOpenSettingsCommand As ICommand

        Public Property StatusText As String
            Get
                Return _statusText
            End Get
            Set(value As String)
                SetProperty(_statusText, If(value, String.Empty))
            End Set
        End Property

        Public Property CurrentThreadText As String
            Get
                Return _currentThreadText
            End Get
            Set(value As String)
                SetProperty(_currentThreadText, If(value, String.Empty))
            End Set
        End Property

        Public Property CurrentTurnText As String
            Get
                Return _currentTurnText
            End Get
            Set(value As String)
                SetProperty(_currentTurnText, If(value, String.Empty))
            End Set
        End Property

        Public Property ConnectionStateText As String
            Get
                Return _connectionStateText
            End Get
            Set(value As String)
                SetProperty(_connectionStateText, If(value, String.Empty))
            End Set
        End Property

        Public Property WorkspaceHintText As String
            Get
                Return _workspaceHintText
            End Get
            Set(value As String)
                SetProperty(_workspaceHintText, If(value, String.Empty))
            End Set
        End Property

        Public Property SidebarNewThreadButtonText As String
            Get
                Return _sidebarNewThreadButtonText
            End Get
            Set(value As String)
                SetProperty(_sidebarNewThreadButtonText, If(value, String.Empty))
            End Set
        End Property

        Public Property SidebarNewThreadToolTip As Object
            Get
                Return _sidebarNewThreadToolTip
            End Get
            Set(value As Object)
                SetProperty(_sidebarNewThreadToolTip, value)
            End Set
        End Property

        Public Property SidebarMainViewVisibility As Visibility
            Get
                Return _sidebarMainViewVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_sidebarMainViewVisibility, value)
            End Set
        End Property

        Public Property SidebarSettingsViewVisibility As Visibility
            Get
                Return _sidebarSettingsViewVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_sidebarSettingsViewVisibility, value)
            End Set
        End Property

        Public Property SidebarNewThreadNavTag As String
            Get
                Return _sidebarNewThreadNavTag
            End Get
            Set(value As String)
                SetProperty(_sidebarNewThreadNavTag, If(value, String.Empty))
            End Set
        End Property

        Public Property SidebarAutomationsNavTag As String
            Get
                Return _sidebarAutomationsNavTag
            End Get
            Set(value As String)
                SetProperty(_sidebarAutomationsNavTag, If(value, String.Empty))
            End Set
        End Property

        Public Property SidebarSkillsNavTag As String
            Get
                Return _sidebarSkillsNavTag
            End Get
            Set(value As String)
                SetProperty(_sidebarSkillsNavTag, If(value, String.Empty))
            End Set
        End Property

        Public Property SidebarSettingsNavTag As String
            Get
                Return _sidebarSettingsNavTag
            End Get
            Set(value As String)
                SetProperty(_sidebarSettingsNavTag, If(value, String.Empty))
            End Set
        End Property

        Public Property IsSidebarNewThreadEnabled As Boolean
            Get
                Return _isSidebarNewThreadEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isSidebarNewThreadEnabled, value)
            End Set
        End Property

        Public Property IsSidebarAutomationsEnabled As Boolean
            Get
                Return _isSidebarAutomationsEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isSidebarAutomationsEnabled, value)
            End Set
        End Property

        Public Property IsSidebarSkillsEnabled As Boolean
            Get
                Return _isSidebarSkillsEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isSidebarSkillsEnabled, value)
            End Set
        End Property

        Public Property IsSidebarSettingsEnabled As Boolean
            Get
                Return _isSidebarSettingsEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isSidebarSettingsEnabled, value)
            End Set
        End Property

        Public Property IsSettingsBackEnabled As Boolean
            Get
                Return _isSettingsBackEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isSettingsBackEnabled, value)
            End Set
        End Property

        Public Property IsQuickOpenVscEnabled As Boolean
            Get
                Return _isQuickOpenVscEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isQuickOpenVscEnabled, value)
            End Set
        End Property

        Public Property IsQuickOpenTerminalEnabled As Boolean
            Get
                Return _isQuickOpenTerminalEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isQuickOpenTerminalEnabled, value)
            End Set
        End Property

        Public Property AreThreadLegacyFilterControlsEnabled As Boolean
            Get
                Return _areThreadLegacyFilterControlsEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_areThreadLegacyFilterControlsEnabled, value)
            End Set
        End Property

        Public Property ShellSendCommand As ICommand
            Get
                Return _shellSendCommand
            End Get
            Set(value As ICommand)
                SetProperty(_shellSendCommand, value)
            End Set
        End Property

        Public Property ShellRefreshThreadsCommand As ICommand
            Get
                Return _shellRefreshThreadsCommand
            End Get
            Set(value As ICommand)
                SetProperty(_shellRefreshThreadsCommand, value)
            End Set
        End Property

        Public Property ShellRefreshModelsCommand As ICommand
            Get
                Return _shellRefreshModelsCommand
            End Get
            Set(value As ICommand)
                SetProperty(_shellRefreshModelsCommand, value)
            End Set
        End Property

        Public Property ShellNewThreadCommand As ICommand
            Get
                Return _shellNewThreadCommand
            End Get
            Set(value As ICommand)
                SetProperty(_shellNewThreadCommand, value)
            End Set
        End Property

        Public Property ShellFocusThreadSearchCommand As ICommand
            Get
                Return _shellFocusThreadSearchCommand
            End Get
            Set(value As ICommand)
                SetProperty(_shellFocusThreadSearchCommand, value)
            End Set
        End Property

        Public Property ShellOpenSettingsCommand As ICommand
            Get
                Return _shellOpenSettingsCommand
            End Get
            Set(value As ICommand)
                SetProperty(_shellOpenSettingsCommand, value)
            End Set
        End Property

        Public ReadOnly Property TurnComposer As TurnComposerViewModel
            Get
                Return _turnComposer
            End Get
        End Property

        Public ReadOnly Property ThreadsPanel As ThreadsPanelViewModel
            Get
                Return _threadsPanel
            End Get
        End Property

        Public ReadOnly Property TranscriptPanel As TranscriptPanelViewModel
            Get
                Return _transcriptPanel
            End Get
        End Property

        Public ReadOnly Property ApprovalPanel As ApprovalPanelViewModel
            Get
                Return _approvalPanel
            End Get
        End Property

        Public ReadOnly Property SettingsPanel As SettingsPanelViewModel
            Get
                Return _settingsPanel
            End Get
        End Property

        Public ReadOnly Property SessionState As SessionStateViewModel
            Get
                Return _sessionState
            End Get
        End Property
    End Class
End Namespace
