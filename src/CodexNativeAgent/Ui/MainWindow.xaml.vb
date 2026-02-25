Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Interop
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Threading
Imports CodexNativeAgent.AppServer
Imports CodexNativeAgent.Services
Imports CodexNativeAgent.Ui.Coordinators
Imports CodexNativeAgent.Ui.Mvvm
Imports CodexNativeAgent.Ui.ViewModels
Imports CodexNativeAgent.Ui.ViewModels.Threads

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Inherits Window

        Private NotInheritable Class ModelListEntry
            Public Property Id As String = String.Empty
            Public Property DisplayName As String = String.Empty
            Public Property IsDefault As Boolean

            Public Overrides Function ToString() As String
                Dim name = If(String.IsNullOrWhiteSpace(DisplayName), Id, DisplayName)
                Return name
            End Function
        End Class

        Private NotInheritable Class PendingUserEcho
            Public Property Text As String = String.Empty
            Public Property AddedUtc As DateTimeOffset
        End Class

        Private Structure ProcessCaptureResult
            Public Property ExitCode As Integer
            Public Property OutputText As String
            Public Property ErrorText As String
        End Structure

        Private NotInheritable Class GitPanelSnapshot
            Public Property WorkingDirectory As String = String.Empty
            Public Property RepoRoot As String = String.Empty
            Public Property RepoName As String = String.Empty
            Public Property BranchName As String = String.Empty
            Public Property StatusSummary As String = String.Empty
            Public Property ChangesText As String = String.Empty
            Public Property CommitsText As String = String.Empty
            Public Property BranchesText As String = String.Empty
            Public Property ChangedFiles As New List(Of GitChangedFileListEntry)()
            Public Property Commits As New List(Of GitCommitListEntry)()
            Public Property Branches As New List(Of GitBranchListEntry)()
            Public Property AddedLineCount As Integer?
            Public Property RemovedLineCount As Integer?
            Public Property ErrorMessage As String = String.Empty
            Public Property LoadedAtLocal As DateTime = DateTime.Now
        End Class

        Private NotInheritable Class GitChangedFileListEntry
            Public Property StatusCode As String = String.Empty
            Public Property FilePath As String = String.Empty
            Public Property DisplayPath As String = String.Empty
            Public Property AddedLineCount As Integer?
            Public Property RemovedLineCount As Integer?
            Public Property IsUntracked As Boolean
            Public Property FileIconSource As ImageSource

            Public ReadOnly Property StatusKind As String
                Get
                    If IsUntracked Then
                        Return "untracked"
                    End If

                    Dim code = If(StatusCode, String.Empty).PadRight(2)
                    Dim x = code(0)
                    Dim y = code(1)

                    If x = "U"c OrElse y = "U"c Then
                        Return "conflict"
                    End If

                    Dim primary = If(x <> " "c, x, y)
                    Select Case primary
                        Case "A"c
                            Return "added"
                        Case "D"c
                            Return "deleted"
                        Case "R"c
                            Return "renamed"
                        Case "C"c
                            Return "copied"
                        Case "M"c, "T"c
                            Return "modified"
                        Case Else
                            Return "other"
                    End Select
                End Get
            End Property

            Public ReadOnly Property StatusBadgeText As String
                Get
                    Dim text = If(StatusCode, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(text) Then
                        Return If(IsUntracked, "U", "--")
                    End If

                    If IsUntracked AndAlso StringComparer.Ordinal.Equals(text, "??") Then
                        Return "U"
                    End If

                    Return text
                End Get
            End Property

            Public ReadOnly Property AddedLinesText As String
                Get
                    If AddedLineCount.HasValue AndAlso AddedLineCount.Value > 0 Then
                        Return $"+{AddedLineCount.Value}"
                    End If

                    Return String.Empty
                End Get
            End Property

            Public ReadOnly Property RemovedLinesText As String
                Get
                    If RemovedLineCount.HasValue AndAlso RemovedLineCount.Value > 0 Then
                        Return $"-{RemovedLineCount.Value}"
                    End If

                    Return String.Empty
                End Get
            End Property

            Public ReadOnly Property AddedLinesVisibility As Visibility
                Get
                    Return If(String.IsNullOrWhiteSpace(AddedLinesText), Visibility.Collapsed, Visibility.Visible)
                End Get
            End Property

            Public ReadOnly Property RemovedLinesVisibility As Visibility
                Get
                    Return If(String.IsNullOrWhiteSpace(RemovedLinesText), Visibility.Collapsed, Visibility.Visible)
                End Get
            End Property

            Public ReadOnly Property DisplayText As String
                Get
                    Dim pathText = If(String.IsNullOrWhiteSpace(DisplayPath), FilePath, DisplayPath)
                    Dim statusText = If(String.IsNullOrWhiteSpace(StatusBadgeText), "--", StatusBadgeText)
                    Dim parts As New List(Of String) From {$"{statusText} {pathText}"}
                    If AddedLineCount.HasValue AndAlso AddedLineCount.Value > 0 Then
                        parts.Add($"+{AddedLineCount.Value}")
                    End If
                    If RemovedLineCount.HasValue AndAlso RemovedLineCount.Value > 0 Then
                        parts.Add($"-{RemovedLineCount.Value}")
                    End If

                    Return String.Join("    ", parts)
                End Get
            End Property

            Public ReadOnly Property FileIconVisibility As Visibility
                Get
                    Return If(FileIconSource Is Nothing, Visibility.Collapsed, Visibility.Visible)
                End Get
            End Property

            Public ReadOnly Property FileIconFallbackVisibility As Visibility
                Get
                    Return If(FileIconSource Is Nothing, Visibility.Visible, Visibility.Collapsed)
                End Get
            End Property

            Public ReadOnly Property DisplayPathPrefixText As String
                Get
                    Dim parts = BuildDisplayPathParts(If(String.IsNullOrWhiteSpace(DisplayPath), FilePath, DisplayPath))
                    Return parts.Prefix
                End Get
            End Property

            Public ReadOnly Property DisplayPathFileNameText As String
                Get
                    Dim parts = BuildDisplayPathParts(If(String.IsNullOrWhiteSpace(DisplayPath), FilePath, DisplayPath))
                    Return parts.FileName
                End Get
            End Property

            Private Shared Function BuildDisplayPathParts(pathText As String) As (Prefix As String, FileName As String)
                Dim text = If(pathText, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(text) Then
                    Return (String.Empty, String.Empty)
                End If

                Dim renamePrefix As String = String.Empty
                Dim targetPath = text
                Dim renameArrowIndex = text.IndexOf(" -> ", StringComparison.Ordinal)
                If renameArrowIndex >= 0 Then
                    renamePrefix = text.Substring(0, renameArrowIndex + 4)
                    targetPath = text.Substring(renameArrowIndex + 4)
                End If

                Dim lastSlash = Math.Max(targetPath.LastIndexOf("/"c), targetPath.LastIndexOf("\"c))
                If lastSlash < 0 Then
                    Return (renamePrefix, CompactFileNamePreservingEnd(targetPath))
                End If

                Dim prefix = renamePrefix & targetPath.Substring(0, lastSlash + 1)
                Dim fileName = targetPath.Substring(lastSlash + 1)
                If String.IsNullOrWhiteSpace(fileName) Then
                    Return (String.Empty, CompactFileNamePreservingEnd(text))
                End If

                Return (prefix, CompactFileNamePreservingEnd(fileName))
            End Function

            Private Shared Function CompactFileNamePreservingEnd(fileName As String) As String
                Dim text = If(fileName, String.Empty)
                If String.IsNullOrWhiteSpace(text) Then
                    Return String.Empty
                End If

                Const maxLen As Integer = 34
                If text.Length <= maxLen Then
                    Return text
                End If

                Dim ext = Path.GetExtension(text)
                Dim stem = text
                If Not String.IsNullOrWhiteSpace(ext) AndAlso ext.Length < text.Length Then
                    stem = text.Substring(0, text.Length - ext.Length)
                Else
                    ext = String.Empty
                End If

                If String.IsNullOrEmpty(stem) Then
                    Return text
                End If

                Dim tailStemLen = Math.Min(8, Math.Max(3, stem.Length \ 3))
                Dim headStemLen = Math.Max(5, maxLen - ext.Length - 3 - tailStemLen)
                If headStemLen + tailStemLen >= stem.Length Then
                    Return text
                End If

                Return stem.Substring(0, headStemLen) &
                       "..." &
                       stem.Substring(stem.Length - tailStemLen) &
                       ext
            End Function
        End Class

        Private NotInheritable Class GitDiffPreviewLineEntry
            Public Property Text As String = String.Empty
            Public Property Kind As String = "context"
        End Class

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Private Structure SHFILEINFO
            Public hIcon As IntPtr
            Public iIcon As Integer
            Public dwAttributes As UInteger
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
            Public szDisplayName As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
            Public szTypeName As String
        End Structure

        <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function SHGetFileInfo(pszPath As String,
                                              dwFileAttributes As UInteger,
                                              ByRef psfi As SHFILEINFO,
                                              cbFileInfo As UInteger,
                                              uFlags As UInteger) As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
        End Function

        Private Const SHGFI_ICON As UInteger = &H100UI
        Private Const SHGFI_SMALLICON As UInteger = &H1UI
        Private Const SHGFI_USEFILEATTRIBUTES As UInteger = &H10UI
        Private Const FILE_ATTRIBUTE_DIRECTORY As UInteger = &H10UI
        Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80UI
        Private Const LeftSidebarDefaultDockWidth As Double = 472.0R
        Private Const LeftSidebarVisibleMinWidth As Double = 368.0R
        Private Const LeftSidebarSplitterVisibleWidth As Double = 8.0R
        Private Const WorkspaceResponsiveMinWidth As Double = 620.0R
        Private Const WorkspaceResponsiveMinWidthWhenGitPaneOpen As Double = 520.0R
        Private Shared ReadOnly _gitFileIconCacheLock As New Object()
        Private Shared ReadOnly _gitFileIconCache As New Dictionary(Of String, ImageSource)(StringComparer.OrdinalIgnoreCase)

        Private NotInheritable Class GitCommitListEntry
            Public Property Sha As String = String.Empty
            Public Property ShortSha As String = String.Empty
            Public Property Subject As String = String.Empty
            Public Property RelativeTime As String = String.Empty
            Public Property Decorations As String = String.Empty

            Public ReadOnly Property DisplayText As String
                Get
                    Dim headline = $"{If(String.IsNullOrWhiteSpace(ShortSha), "???????", ShortSha)}  {If(Subject, String.Empty)}".Trim()
                    If String.IsNullOrWhiteSpace(RelativeTime) Then
                        Return headline
                    End If

                    Return $"{headline}    ({RelativeTime})"
                End Get
            End Property
        End Class

        Private NotInheritable Class GitBranchListEntry
            Public Property Name As String = String.Empty
            Public Property IsCurrent As Boolean
            Public Property IsRemote As Boolean
            Public Property CommitShortSha As String = String.Empty
            Public Property RelativeTime As String = String.Empty
            Public Property Subject As String = String.Empty

            Public ReadOnly Property DisplayText As String
                Get
                    Dim prefix = If(IsCurrent, "● ", If(IsRemote, "◌ ", "○ "))
                    Dim headline = prefix & If(Name, String.Empty)
                    If Not String.IsNullOrWhiteSpace(CommitShortSha) Then
                        headline &= $"  {CommitShortSha}"
                    End If

                    If Not String.IsNullOrWhiteSpace(RelativeTime) Then
                        headline &= $"  ({RelativeTime})"
                    End If

                    Return headline
                End Get
            End Property
        End Class


        Private NotInheritable Class AppSettings
            Public Property CodexPath As String = String.Empty
            Public Property ServerArgs As String = "app-server"
            Public Property WorkingDir As String = String.Empty
            Public Property WindowsCodexHome As String = String.Empty
            Public Property RememberApiKey As Boolean
            Public Property AutoLoginApiKey As Boolean
            Public Property AutoReconnect As Boolean = True
            Public Property DisableWorkspaceHintOverlay As Boolean
            Public Property DisableConnectionInitializedToast As Boolean
            Public Property DisableThreadsPanelHints As Boolean
            Public Property ShowEventDotsInTranscript As Boolean
            Public Property ShowSystemDotsInTranscript As Boolean
            Public Property FilterThreadsByWorkingDir As Boolean
            Public Property EncryptedApiKey As String = String.Empty
            Public Property ThemeMode As String = AppAppearanceManager.LightTheme
            Public Property DensityMode As String = AppAppearanceManager.ComfortableDensity
            Public Property TurnComposerPickersCollapsed As Boolean
        End Class

        Private ReadOnly _accountService As IAccountService
        Private ReadOnly _connectionService As IConnectionService
        Private ReadOnly _approvalService As IApprovalService
        Private ReadOnly _threadService As IThreadService
        Private ReadOnly _turnService As ITurnService
        Private ReadOnly _settingsJsonOptions As New JsonSerializerOptions With {
            .WriteIndented = True
        }
        Private ReadOnly _settingsFilePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CodexNativeAgent", "settings.json")
        Private ReadOnly _toastTimer As New DispatcherTimer()
        Private ReadOnly _watchdogTimer As New DispatcherTimer()
        Private ReadOnly _reconnectUiTimer As New DispatcherTimer()
        Private ReadOnly _workspaceHintOverlayTimer As New DispatcherTimer()
        Private ReadOnly _threadEntries As New List(Of ThreadListEntry)()
        Private ReadOnly _expandedThreadProjectGroups As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _pendingLocalUserEchoes As New Queue(Of PendingUserEcho)()
        Private ReadOnly _suppressedServerUserEchoItemKeys As New HashSet(Of String)(StringComparer.Ordinal)
        Private ReadOnly _snapshotAssistantPhaseHintsByItemKey As New Dictionary(Of String, String)(StringComparer.Ordinal)
        Private ReadOnly _snapshotAssistantPhaseHintsLock As New Object()
        Private Shared ReadOnly PendingUserEchoMaxAge As TimeSpan = TimeSpan.FromSeconds(30)

        Private _client As CodexAppServerClient
        Private _currentThreadId As String = String.Empty
        Private _currentTurnId As String = String.Empty
        Private _notificationRuntimeThreadId As String = String.Empty
        Private _notificationRuntimeTurnId As String = String.Empty
        Private _currentLoginId As String = String.Empty
        Private _disconnecting As Boolean
        Private _threadsLoading As Boolean
        Private _threadSelectionLoadCts As CancellationTokenSource
        Private _threadContentLoading As Boolean
        Private _suppressThreadSelectionEvents As Boolean
        Private _suppressThreadToolbarMenuEvents As Boolean
        Private _threadContextTarget As ThreadListEntry
        Private _threadGroupContextTarget As ThreadGroupHeaderEntry
        Private _currentThreadCwd As String = String.Empty
        Private _newThreadTargetOverrideCwd As String = String.Empty
        Private _connectionExpected As Boolean
        Private _reconnectInProgress As Boolean
        Private _reconnectAttempt As Integer
        Private _reconnectCts As CancellationTokenSource
        Private _lastActivityUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Private _lastWatchdogWarningUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _nextReconnectAttemptUtc As DateTimeOffset?
        Private _isAuthenticated As Boolean
        Private _authRequiredNoticeShown As Boolean
        Private _startupConnectAttempted As Boolean
        Private _modelsLoadedAtLeastOnce As Boolean
        Private _threadsLoadedAtLeastOnce As Boolean
        Private _workspaceBootstrapInProgress As Boolean
        Private _currentTheme As String = AppAppearanceManager.LightTheme
        Private _currentDensity As String = AppAppearanceManager.ComfortableDensity
        Private _suppressAppearanceUiChange As Boolean
        Private _turnComposerPickersCollapsed As Boolean
        Private _turnComposerPickersExpandedWidth As Double = 434.0R
        Private _transcriptAutoScrollEnabled As Boolean = True
        Private _suppressTranscriptScrollTracking As Boolean
        Private _pendingNewThreadFirstPromptSelection As Boolean
        Private _threadsPanelHintBubbleHintKey As String = String.Empty
        Private _threadsPanelHintBubbleDismissedForCurrentHint As Boolean
        Private _workspaceHintOverlayHintKey As String = String.Empty
        Private _workspaceHintOverlayDismissedForCurrentHint As Boolean
        Private _gitPanelLoadVersion As Integer
        Private _gitPanelDiffPreviewLoadVersion As Integer
        Private _gitPanelCommitPreviewLoadVersion As Integer
        Private _gitPanelBranchPreviewLoadVersion As Integer
        Private _suppressGitPanelSelectionEvents As Boolean
        Private _gitPanelActiveTab As String = "changes"
        Private _gitPanelDockWidth As Double = 560.0R
        Private _leftSidebarDockWidth As Double = LeftSidebarDefaultDockWidth
        Private _isLeftSidebarVisible As Boolean = True
        Private _currentGitPanelSnapshot As GitPanelSnapshot
        Private _gitPanelSelectedDiffFilePath As String = String.Empty
        Private _protocolDialogWindow As Window
        Private _isStatusBarExpanded As Boolean
        Private _settings As New AppSettings()
        Private ReadOnly _viewModel As New MainWindowViewModel()
        Private ReadOnly _sessionCoordinator As SessionCoordinator
        Private ReadOnly _sessionNotificationCoordinator As SessionNotificationCoordinator
        Private ReadOnly _threadLiveSessionRegistry As New ThreadLiveSessionRegistry()
        Private ReadOnly _threadWorkflowCoordinator As ThreadWorkflowCoordinator
        Private ReadOnly _turnWorkflowCoordinator As TurnWorkflowCoordinator
        Private ReadOnly _shellCommandCoordinator As ShellCommandCoordinator
        Private ReadOnly _settingsStore As IAppSettingsStore

        Public Sub New()
            InitializeComponent()
            DataContext = _viewModel
            _accountService = New CodexAccountService(Function() CurrentClient())
            _connectionService = New CodexConnectionService()
            _approvalService = New CodexApprovalService()
            _threadService = New CodexThreadService(Function() CurrentClient())
            _turnService = New CodexTurnService(Function() CurrentClient())
            _sessionCoordinator = New SessionCoordinator(
                _viewModel,
                Function(operation) RunUiActionAsync(operation),
                AddressOf ConnectAsync,
                AddressOf DisconnectAsync,
                AddressOf ReconnectNowAsync,
                AddressOf RefreshAuthenticationGateAsync,
                AddressOf LoginApiKeyAsync,
                AddressOf LoginChatGptAsync,
                AddressOf CancelLoginAsync,
                AddressOf LogoutAsync,
                AddressOf ReadRateLimitsAsync,
                AddressOf LoginExternalTokensAsync)
            _sessionNotificationCoordinator = New SessionNotificationCoordinator()
            _threadWorkflowCoordinator = New ThreadWorkflowCoordinator()
            _turnWorkflowCoordinator = New TurnWorkflowCoordinator(
                _viewModel,
                _turnService,
                _approvalService,
                Function(operation) RunUiActionAsync(operation))
            AddHandler _turnWorkflowCoordinator.ApprovalResolved,
                AddressOf OnTurnWorkflowApprovalResolved
            _shellCommandCoordinator = New ShellCommandCoordinator(
                _viewModel,
                Function(operation) RunUiActionAsync(operation),
                AddressOf FireAndForget)
            _settingsStore = New JsonAppSettingsStore(_settingsFilePath, _settingsJsonOptions)
            AddHandler _viewModel.TranscriptPanel.Items.CollectionChanged,
                Sub(sender, e)
                    UpdateWorkspaceEmptyStateVisibility()
                End Sub
            AddHandler _viewModel.ThreadsPanel.PropertyChanged,
                Sub(sender, e)
                    If e Is Nothing OrElse
                       Not StringComparer.Ordinal.Equals(e.PropertyName, NameOf(ThreadsPanelViewModel.StateText)) Then
                        Return
                    End If

                    UpdateThreadsPanelStateHintBubbleVisibility()
                End Sub
            InitializeSessionCoordinatorBindings()
            _turnWorkflowCoordinator.BindCommands(AddressOf StartTurnAsync,
                                                  AddressOf SteerTurnAsync,
                                                  AddressOf InterruptTurnAsync,
                                                  AddressOf ResolveApprovalAsync)
            InitializeMvvmCommandBindings()

            InitializeUiDefaults()
            InitializeEventHandlers()
            InitializeStatusUi()
            InitializeReliabilityLayer()
            InitializeDefaults()
            ShowWorkspaceView()
            RefreshControlStates()
            ShowStatus("Ready.")
        End Sub

        Private Sub InitializeMvvmCommandBindings()
            If SidebarPaneHost.ThreadSortContextMenu IsNot Nothing Then
                SidebarPaneHost.ThreadSortContextMenu.DataContext = _viewModel
            End If

            If SidebarPaneHost.ThreadFilterContextMenu IsNot Nothing Then
                SidebarPaneHost.ThreadFilterContextMenu.DataContext = _viewModel
            End If

            If SidebarPaneHost.ThreadItemContextMenu IsNot Nothing Then
                SidebarPaneHost.ThreadItemContextMenu.DataContext = _viewModel
            End If

            _shellCommandCoordinator.BindCommands(
                AddressOf StartTurnAsync,
                AddressOf RefreshThreadsAsync,
                AddressOf RefreshModelsAsync,
                AddressOf StartThreadAsync,
                Sub()
                    ShowThreadsSidebarTab()
                    SidebarPaneHost.TxtThreadSearch.Focus()
                    SidebarPaneHost.TxtThreadSearch.SelectAll()
                End Sub,
                AddressOf ShowControlCenterTab,
                Sub() OpenThreadToolbarMenu(SidebarPaneHost.BtnThreadSortMenu, SidebarPaneHost.ThreadSortContextMenu),
                Sub() OpenThreadToolbarMenu(SidebarPaneHost.BtnThreadFilterMenu, SidebarPaneHost.ThreadFilterContextMenu),
                Sub(targetIndex)
                    If targetIndex < 0 Then
                        Return
                    End If

                    ApplyThreadSortMenuSelection(targetIndex)
                End Sub,
                AddressOf ApplyThreadFilterMenuToggle,
                AddressOf SelectThreadFromContextMenu,
                AddressOf RefreshThreadFromContextMenuAsync,
                AddressOf ForkThreadFromContextMenuAsync,
                AddressOf ArchiveThreadFromContextMenuAsync,
                AddressOf UnarchiveThreadFromContextMenuAsync,
                AddressOf StartThreadFromGroupHeaderContextMenuAsync,
                AddressOf OpenThreadGroupInVsCodeFromContextMenuAsync,
                AddressOf OpenThreadGroupInTerminalFromContextMenuAsync,
                AddressOf ToggleTheme,
                AddressOf ExportDiagnosticsAsync)

            _sessionCoordinator.BindSettingsCommands()
        End Sub

        Private Sub InitializeUiDefaults()
            If _viewModel.ThreadsPanel.SortIndex < 0 Then
                _viewModel.ThreadsPanel.SortIndex = 0
            End If

            If WorkspacePaneHost.CmbReasoningEffort.SelectedIndex < 0 Then
                WorkspacePaneHost.CmbReasoningEffort.SelectedIndex = 2
            End If

            If StatusBarPaneHost.CmbApprovalPolicy.SelectedIndex < 0 Then
                StatusBarPaneHost.CmbApprovalPolicy.SelectedIndex = 0
            End If

            If StatusBarPaneHost.CmbSandbox.SelectedIndex < 0 Then
                StatusBarPaneHost.CmbSandbox.SelectedIndex = 0
            End If

            If _viewModel.SettingsPanel.DensityIndex < 0 Then
                _viewModel.SettingsPanel.DensityIndex = 0
            End If

            If SidebarPaneHost.CmbExternalPlanType.Items.Count = 0 Then
                SidebarPaneHost.CmbExternalPlanType.Items.Add("")
                SidebarPaneHost.CmbExternalPlanType.Items.Add("free")
                SidebarPaneHost.CmbExternalPlanType.Items.Add("go")
                SidebarPaneHost.CmbExternalPlanType.Items.Add("plus")
                SidebarPaneHost.CmbExternalPlanType.Items.Add("pro")
                SidebarPaneHost.CmbExternalPlanType.Items.Add("team")
                SidebarPaneHost.CmbExternalPlanType.Items.Add("business")
                SidebarPaneHost.CmbExternalPlanType.Items.Add("enterprise")
                SidebarPaneHost.CmbExternalPlanType.Items.Add("edu")
                SidebarPaneHost.CmbExternalPlanType.Items.Add("unknown")
            End If
            SidebarPaneHost.CmbExternalPlanType.SelectedIndex = 0
            _viewModel.SettingsPanel.ExternalPlanType = String.Empty

            _viewModel.ApprovalPanel.ResetLifecycleState()
            _viewModel.SettingsPanel.RateLimitsText = "No rate-limit data loaded yet."
            _viewModel.ThreadsPanel.StateText = "No threads loaded yet."
            _viewModel.CurrentThreadText = "New thread"
            _viewModel.CurrentTurnText = "Turn: 0"
            UpdateConnectionStateTextFromSession(syncFirst:=False)
            _viewModel.SettingsPanel.ReconnectCountdownText = "Reconnect: not scheduled."
            SetTranscriptLoadingState(False)
            UpdateSidebarSelectionState(showSettings:=False)
            ApplyTurnComposerPickersCollapsedState(animated:=False, persist:=False)
            SyncAppearanceControls()
            SyncThreadToolbarMenus()
            SyncNewThreadTargetChip()
            SetStatusBarExpanded(False)
            SyncThreadsSidebarToggleVisual()
        End Sub

        Private Sub InitializeEventHandlers()
            AddHandler Me.SizeChanged, Sub(sender, e) UpdateMainPaneResizeBounds()
            AddHandler BtnStatusBarToggle.Click, Sub(sender, e) ToggleStatusBarVisibility()
            AddHandler WorkspacePaneHost.BtnQuickToggleThreadsPanel.Click, Sub(sender, e) ToggleThreadsSidebarVisibility()
            AddHandler SidebarPaneHost.BtnSidebarNewThread.Click, Async Sub(sender, e)
                                                       ShowWorkspaceView()
                                                       Await RunUiActionAsync(AddressOf StartThreadAsync)
                                                   End Sub
            AddHandler SidebarPaneHost.SidebarNewThreadContextMenu.Opened, Sub(sender, e) SyncSidebarNewThreadMenu()
            AddHandler SidebarPaneHost.MnuSidebarNewThreadChooseFolder.Click, Async Sub(sender, e)
                                                               ShowWorkspaceView()
                                                               Await RunUiActionAsync(AddressOf ChooseFolderAndStartNewThreadAsync)
                                                           End Sub
            AddHandler SidebarPaneHost.BtnSidebarSettings.Click, Sub(sender, e) ShowSettingsView()
            AddHandler SidebarPaneHost.BtnSettingsBack.Click, Sub(sender, e) ShowWorkspaceView()
            AddHandler SidebarPaneHost.CmbDensity.SelectionChanged, Sub(sender, e) OnDensitySelectionChanged()
            AddHandler SidebarPaneHost.TxtWorkingDir.TextChanged, Sub(sender, e) SyncNewThreadTargetChip()

            AddHandler SidebarPaneHost.ChkAutoReconnect.Checked, Sub(sender, e) SaveSettings()
            AddHandler SidebarPaneHost.ChkAutoReconnect.Unchecked, Sub(sender, e) SaveSettings()
            AddHandler SidebarPaneHost.ChkDisableWorkspaceHintOverlay.Checked,
                Sub(sender, e)
                    SaveSettings()
                    DismissWorkspaceHintOverlay(dismissCurrentHint:=False)
                    UpdateWorkspaceHintOverlayVisibility()
                End Sub
            AddHandler SidebarPaneHost.ChkDisableWorkspaceHintOverlay.Unchecked,
                Sub(sender, e)
                    SaveSettings()
                    _workspaceHintOverlayDismissedForCurrentHint = False
                    UpdateWorkspaceHintOverlayVisibility()
                End Sub
            AddHandler SidebarPaneHost.ChkDisableConnectionInitializedToast.Checked, Sub(sender, e) SaveSettings()
            AddHandler SidebarPaneHost.ChkDisableConnectionInitializedToast.Unchecked, Sub(sender, e) SaveSettings()
            AddHandler SidebarPaneHost.ChkDisableThreadsPanelHints.Checked,
                Sub(sender, e)
                    SaveSettings()
                    DismissThreadsPanelStateHint(dismissCurrentHint:=False)
                    UpdateThreadsPanelStateHintBubbleVisibility()
                End Sub
            AddHandler SidebarPaneHost.ChkDisableThreadsPanelHints.Unchecked,
                Sub(sender, e)
                    SaveSettings()
                    _threadsPanelHintBubbleDismissedForCurrentHint = False
                    UpdateThreadsPanelStateHintBubbleVisibility()
                End Sub
            AddHandler SidebarPaneHost.ChkShowEventDotsInTranscript.Checked,
                Sub(sender, e)
                    SaveSettings()
                    ApplyTranscriptTimelineDotVisibilitySettings()
                End Sub
            AddHandler SidebarPaneHost.ChkShowEventDotsInTranscript.Unchecked,
                Sub(sender, e)
                    SaveSettings()
                    ApplyTranscriptTimelineDotVisibilitySettings()
                End Sub
            AddHandler SidebarPaneHost.ChkShowSystemDotsInTranscript.Checked,
                Sub(sender, e)
                    SaveSettings()
                    ApplyTranscriptTimelineDotVisibilitySettings()
                End Sub
            AddHandler SidebarPaneHost.ChkShowSystemDotsInTranscript.Unchecked,
                Sub(sender, e)
                    SaveSettings()
                    ApplyTranscriptTimelineDotVisibilitySettings()
                End Sub

            AddHandler SidebarPaneHost.ChkRememberApiKey.Checked, Sub(sender, e) SaveSettings()
            AddHandler SidebarPaneHost.ChkRememberApiKey.Unchecked, Sub(sender, e) SaveSettings()
            AddHandler SidebarPaneHost.ChkAutoLoginApiKey.Checked, Sub(sender, e) SaveSettings()
            AddHandler SidebarPaneHost.ChkAutoLoginApiKey.Unchecked, Sub(sender, e) SaveSettings()

            AddHandler SidebarPaneHost.ChkShowArchivedThreads.Checked, Async Sub(sender, e) Await RunUiActionAsync(AddressOf RefreshThreadsAsync)
            AddHandler SidebarPaneHost.ChkShowArchivedThreads.Unchecked, Async Sub(sender, e) Await RunUiActionAsync(AddressOf RefreshThreadsAsync)
            AddHandler SidebarPaneHost.ChkFilterThreadsByWorkingDir.Checked, Async Sub(sender, e) Await RunUiActionAsync(AddressOf RefreshThreadsAsync)
            AddHandler SidebarPaneHost.ChkFilterThreadsByWorkingDir.Unchecked, Async Sub(sender, e) Await RunUiActionAsync(AddressOf RefreshThreadsAsync)
            AddHandler SidebarPaneHost.TxtThreadSearch.TextChanged, Sub(sender, e) ApplyThreadFiltersAndSort()
            AddHandler SidebarPaneHost.CmbThreadSort.SelectionChanged,
                Sub(sender, e)
                    ApplyThreadFiltersAndSort()
                    SyncThreadToolbarMenus()
                End Sub
            AddHandler SidebarPaneHost.ThreadSortContextMenu.Opened, Sub(sender, e) SyncThreadToolbarMenus()
            AddHandler SidebarPaneHost.ThreadFilterContextMenu.Opened, Sub(sender, e) SyncThreadToolbarMenus()
            AddHandler SidebarPaneHost.LstThreads.PreviewMouseLeftButtonDown, AddressOf OnThreadsPreviewMouseLeftButtonDown
            AddHandler SidebarPaneHost.LstThreads.PreviewMouseRightButtonDown, AddressOf OnThreadsPreviewMouseRightButtonDown
            AddHandler SidebarPaneHost.LstThreads.ContextMenuOpening, AddressOf OnThreadsContextMenuOpening
            AddHandler SidebarPaneHost.ThreadItemContextMenu.Closed,
                Sub(sender, e)
                    _threadContextTarget = Nothing
                    _threadGroupContextTarget = Nothing
                End Sub
            AddHandler SidebarPaneHost.LstThreads.SelectionChanged,
                Sub(sender, e)
                    If _suppressThreadSelectionEvents Then
                        Return
                    End If

                    Dim selectedHeader = TryCast(_viewModel.ThreadsPanel.SelectedListItem, ThreadGroupHeaderEntry)
                    If selectedHeader IsNot Nothing Then
                        _suppressThreadSelectionEvents = True
                        _viewModel.ThreadsPanel.SelectedListItem = Nothing
                        _suppressThreadSelectionEvents = False
                        ToggleThreadProjectGroupExpansion(selectedHeader.GroupKey)
                        ApplyThreadFiltersAndSort()
                        Return
                    End If

                    Dim selected = TryCast(_viewModel.ThreadsPanel.SelectedListItem, ThreadListEntry)
                    If selected Is Nothing Then
                        CancelActiveThreadSelectionLoad()
                        ResetThreadSelectionLoadUiState(hideTranscriptLoader:=True)
                    Else
                        FireAndForget(AutoLoadThreadSelectionAsync(selected))
                    End If

                    RefreshControlStates()
                End Sub

            AddHandler WorkspacePaneHost.BtnQuickOpenVsc.Click, Sub(sender, e) OpenWorkspaceInVsCode()
            AddHandler WorkspacePaneHost.BtnQuickOpenGit.Click, Sub(sender, e) ToggleGitPanel()
            AddHandler WorkspacePaneHost.BtnQuickOpenTerminal.Click, Sub(sender, e) OpenWorkspaceInPowerShell()
            AddHandler WorkspacePaneHost.BtnTurnComposerPickersToggle.Click, Sub(sender, e) ToggleTurnComposerPickersCollapsed()
            AddHandler GitPaneHost.BtnGitPanelRefresh.Click, Sub(sender, e) FireAndForget(RefreshGitPanelAsync())
            AddHandler GitPaneHost.BtnGitPanelClose.Click, Sub(sender, e) CloseGitPanel()
            AddHandler GitPaneHost.BtnGitTabChanges.Click, Sub(sender, e) ShowGitPanelTab("changes")
            AddHandler GitPaneHost.BtnGitTabHistory.Click, Sub(sender, e) ShowGitPanelTab("history")
            AddHandler GitPaneHost.BtnGitTabBranches.Click, Sub(sender, e) ShowGitPanelTab("branches")
            AddHandler GitPaneHost.LstGitPanelChanges.SelectionChanged, AddressOf OnGitPanelChangesSelectionChanged
            AddHandler GitPaneHost.LstGitPanelChanges.MouseDoubleClick, AddressOf OnGitPanelChangesMouseDoubleClick
            GitPaneHost.LstGitPanelChanges.AddHandler(Button.ClickEvent,
                                                            New RoutedEventHandler(AddressOf OnGitPanelChangesListButtonClick),
                                                            True)
            AddHandler GitPaneHost.LstGitPanelCommits.SelectionChanged, AddressOf OnGitPanelCommitsSelectionChanged
            AddHandler GitPaneHost.LstGitPanelBranches.SelectionChanged, AddressOf OnGitPanelBranchesSelectionChanged
            InitializeGitPanelUi()
            If WorkspacePaneHost.LstTranscript IsNot Nothing Then
                WorkspacePaneHost.LstTranscript.AddHandler(ScrollViewer.ScrollChangedEvent,
                                                          New ScrollChangedEventHandler(AddressOf OnTranscriptScrollChanged))
            End If
            AddHandler WorkspacePaneHost.BtnDismissWorkspaceHintOverlay.Click,
                Sub(sender, e)
                    DismissWorkspaceHintOverlay()
                End Sub
            AddHandler SidebarPaneHost.BtnDismissThreadsStateHint.Click,
                Sub(sender, e)
                    DismissThreadsPanelStateHint()
                End Sub
            AddHandler BtnToastClose.Click, Sub(sender, e) HideToast()
            AddHandler SidebarPaneHost.BtnSettingsOpenProtocolDialog.Click, Sub(sender, e) ShowProtocolDialog()
        End Sub

        Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
            UpdateMainPaneResizeBounds()
            SyncStatusBarToggleVisual()

            If _startupConnectAttempted Then
                Return
            End If

            _startupConnectAttempted = True
            FireAndForget(RunUiActionAsync(AddressOf AutoConnectOnStartupAsync))
        End Sub

        Private Sub ToggleStatusBarVisibility()
            SetStatusBarExpanded(Not _isStatusBarExpanded)
        End Sub

        Private Sub SetStatusBarExpanded(isExpanded As Boolean)
            _isStatusBarExpanded = isExpanded
            If StatusBarPaneHost IsNot Nothing Then
                StatusBarPaneHost.Visibility = If(isExpanded, Visibility.Visible, Visibility.Collapsed)
            End If

            SyncStatusBarToggleVisual()
        End Sub

        Private Sub SyncStatusBarToggleVisual()
            If TxtStatusBarToggleGlyph IsNot Nothing Then
                ' Up when collapsed (expand), down when expanded (collapse).
                TxtStatusBarToggleGlyph.Text = If(_isStatusBarExpanded, ChrW(&HE70D), ChrW(&HE70E))
            End If

            If BtnStatusBarToggle IsNot Nothing Then
                BtnStatusBarToggle.ToolTip = If(_isStatusBarExpanded, "Hide status bar", "Show status bar")
            End If
        End Sub

        Private Sub ToggleThreadsSidebarVisibility()
            SetThreadsSidebarVisible(Not _isLeftSidebarVisible)
        End Sub

        Private Sub SetThreadsSidebarVisible(isVisible As Boolean)
            If LeftSidebarColumn Is Nothing Then
                _isLeftSidebarVisible = isVisible
                SyncThreadsSidebarToggleVisual()
                Return
            End If

            If isVisible = _isLeftSidebarVisible AndAlso LeftSidebarColumn.ActualWidth >= 1.0R Then
                SyncThreadsSidebarToggleVisual()
                Return
            End If

            If Not isVisible Then
                Dim actualWidth = LeftSidebarColumn.ActualWidth
                If actualWidth > 1.0R Then
                    _leftSidebarDockWidth = Math.Max(LeftSidebarVisibleMinWidth, actualWidth)
                End If
            End If

            _isLeftSidebarVisible = isVisible
            ApplyThreadsSidebarDockState()
            UpdateMainPaneResizeBounds()

            If isVisible AndAlso LeftSidebarColumn IsNot Nothing Then
                Dim targetWidth = Math.Max(LeftSidebarVisibleMinWidth, _leftSidebarDockWidth)
                Dim maxWidth = LeftSidebarColumn.MaxWidth
                If Not Double.IsNaN(maxWidth) AndAlso Not Double.IsInfinity(maxWidth) AndAlso maxWidth > 0 Then
                    targetWidth = Math.Min(targetWidth, maxWidth)
                End If

                LeftSidebarColumn.Width = New GridLength(targetWidth, GridUnitType.Pixel)
            End If
        End Sub

        Private Sub ApplyThreadsSidebarDockState()
            If LeftSidebarColumn IsNot Nothing Then
                LeftSidebarColumn.MinWidth = If(_isLeftSidebarVisible, LeftSidebarVisibleMinWidth, 0)
                If Not _isLeftSidebarVisible Then
                    LeftSidebarColumn.Width = New GridLength(0, GridUnitType.Pixel)
                End If
            End If

            If LeftSidebarSplitterColumn IsNot Nothing Then
                LeftSidebarSplitterColumn.Width = New GridLength(If(_isLeftSidebarVisible, LeftSidebarSplitterVisibleWidth, 0), GridUnitType.Pixel)
            End If

            If LeftSidebarSplitter IsNot Nothing Then
                LeftSidebarSplitter.Visibility = If(_isLeftSidebarVisible, Visibility.Visible, Visibility.Collapsed)
            End If

            SyncThreadsSidebarToggleVisual()
        End Sub

        Private Sub SyncThreadsSidebarToggleVisual()
            If WorkspacePaneHost Is Nothing Then
                Return
            End If

            If WorkspacePaneHost.BtnQuickToggleThreadsPanel IsNot Nothing Then
                WorkspacePaneHost.BtnQuickToggleThreadsPanel.ToolTip = If(_isLeftSidebarVisible, "Hide threads panel", "Show threads panel")
            End If

            If WorkspacePaneHost.BrushThreadsPanelToggleIconMask IsNot Nothing Then
                Dim iconUri = If(_isLeftSidebarVisible,
                                 "pack://application:,,,/Assets/hide-threads-white.png",
                                 "pack://application:,,,/Assets/show-threads-white.png")
                WorkspacePaneHost.BrushThreadsPanelToggleIconMask.ImageSource = New BitmapImage(New Uri(iconUri, UriKind.Absolute))
            End If
        End Sub

        Private Sub UpdateMainPaneResizeBounds()
            If MainSurfaceGrid Is Nothing Then
                Return
            End If

            Dim surfaceWidth = MainSurfaceGrid.ActualWidth
            If Double.IsNaN(surfaceWidth) OrElse Double.IsInfinity(surfaceWidth) OrElse surfaceWidth <= 0 Then
                Return
            End If

            Dim halfSurfaceWidth = Math.Floor(surfaceWidth / 2.0R)
            If halfSurfaceWidth <= 0 Then
                Return
            End If

            Dim isGitPaneVisible =
                RightGitPaneShell IsNot Nothing AndAlso
                RightGitPaneShell.Visibility = Visibility.Visible
            Dim leftSplitterWidth = If(_isLeftSidebarVisible, LeftSidebarSplitterVisibleWidth, 0.0R)
            Dim rightSplitterWidth = If(isGitPaneVisible, 8.0R, 0.0R)
            Dim reservedWorkspaceWidth =
                If(isGitPaneVisible, WorkspaceResponsiveMinWidthWhenGitPaneOpen, WorkspaceResponsiveMinWidth)
            Dim maxCombinedSideWidth = Math.Max(0.0R, surfaceWidth - leftSplitterWidth - rightSplitterWidth - reservedWorkspaceWidth)

            If LeftSidebarColumn IsNot Nothing Then
                If _isLeftSidebarVisible Then
                    LeftSidebarColumn.MinWidth = LeftSidebarVisibleMinWidth
                    LeftSidebarColumn.MaxWidth = Math.Max(LeftSidebarColumn.MinWidth, halfSurfaceWidth)

                    Dim leftActualWidth = LeftSidebarColumn.ActualWidth
                    If leftActualWidth > LeftSidebarColumn.MaxWidth + 0.5R Then
                        LeftSidebarColumn.Width = New GridLength(LeftSidebarColumn.MaxWidth, GridUnitType.Pixel)
                    End If
                Else
                    LeftSidebarColumn.MinWidth = 0
                    LeftSidebarColumn.MaxWidth = Math.Max(0, halfSurfaceWidth)
                    If LeftSidebarColumn.Width.Value > 0.5R OrElse LeftSidebarColumn.ActualWidth > 0.5R Then
                        LeftSidebarColumn.Width = New GridLength(0, GridUnitType.Pixel)
                    End If
                End If
            End If

            If RightGitPaneColumn IsNot Nothing Then
                RightGitPaneColumn.MaxWidth = Math.Max(RightGitPaneColumn.MinWidth, halfSurfaceWidth)

                Dim rightActualWidth = RightGitPaneColumn.ActualWidth
                If rightActualWidth > RightGitPaneColumn.MaxWidth + 0.5R Then
                    RightGitPaneColumn.Width = New GridLength(RightGitPaneColumn.MaxWidth, GridUnitType.Pixel)
                    _gitPanelDockWidth = Math.Min(_gitPanelDockWidth, RightGitPaneColumn.MaxWidth)
                End If
            End If

            ' Reserve breathing room for the center workspace on narrow/portrait windows.
            Dim leftSideWidth = 0.0R
            Dim leftMinWidth = 0.0R
            If LeftSidebarColumn IsNot Nothing AndAlso _isLeftSidebarVisible Then
                leftSideWidth = Math.Max(LeftSidebarColumn.ActualWidth, LeftSidebarColumn.Width.Value)
                leftMinWidth = Math.Max(0.0R, LeftSidebarColumn.MinWidth)
            End If

            Dim rightSideWidth = 0.0R
            Dim rightMinWidth = 0.0R
            If RightGitPaneColumn IsNot Nothing AndAlso isGitPaneVisible Then
                rightSideWidth = Math.Max(RightGitPaneColumn.ActualWidth, RightGitPaneColumn.Width.Value)
                rightMinWidth = Math.Max(0.0R, RightGitPaneColumn.MinWidth)
            End If

            Dim overflow = (leftSideWidth + rightSideWidth) - maxCombinedSideWidth
            If overflow > 0.5R Then
                If LeftSidebarColumn IsNot Nothing AndAlso _isLeftSidebarVisible Then
                    Dim shrinkableLeft = Math.Max(0.0R, leftSideWidth - leftMinWidth)
                    Dim shrinkLeft = Math.Min(overflow, shrinkableLeft)
                    If shrinkLeft > 0.5R Then
                        leftSideWidth = Math.Max(leftMinWidth, leftSideWidth - shrinkLeft)
                        LeftSidebarColumn.Width = New GridLength(leftSideWidth, GridUnitType.Pixel)
                        overflow -= shrinkLeft
                    End If
                End If

                If overflow > 0.5R AndAlso RightGitPaneColumn IsNot Nothing AndAlso isGitPaneVisible Then
                    Dim shrinkableRight = Math.Max(0.0R, rightSideWidth - rightMinWidth)
                    Dim shrinkRight = Math.Min(overflow, shrinkableRight)
                    If shrinkRight > 0.5R Then
                        rightSideWidth = Math.Max(rightMinWidth, rightSideWidth - shrinkRight)
                        RightGitPaneColumn.Width = New GridLength(rightSideWidth, GridUnitType.Pixel)
                        _gitPanelDockWidth = Math.Min(_gitPanelDockWidth, rightSideWidth)
                        overflow -= shrinkRight
                    End If
                End If
            End If
        End Sub

        Private Sub ShowProtocolDialog()
            If _protocolDialogWindow IsNot Nothing AndAlso _protocolDialogWindow.IsVisible Then
                If _protocolDialogWindow.WindowState = WindowState.Minimized Then
                    _protocolDialogWindow.WindowState = WindowState.Normal
                End If

                _protocolDialogWindow.Activate()
                Return
            End If

            Dim closeButton As New Button() With {
                .Content = "Close",
                .Style = TryCast(TryFindResource("ButtonBaseStyle"), Style),
                .Width = 90,
                .HorizontalAlignment = HorizontalAlignment.Right
            }

            Dim protocolViewer As New TextBox() With {
                .IsReadOnly = True,
                .AcceptsReturn = True,
                .TextWrapping = TextWrapping.Wrap,
                .VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                .HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                .Background = Brushes.Transparent,
                .BorderThickness = New Thickness(0),
                .FontFamily = New FontFamily("Cascadia Code")
            }
            protocolViewer.SetBinding(TextBox.TextProperty, New Binding("TranscriptPanel.ProtocolText") With {
                .Source = _viewModel,
                .Mode = BindingMode.OneWay,
                .UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                .FallbackValue = String.Empty,
                .TargetNullValue = String.Empty
            })

            Dim keepScrolledToBottom As TextChangedEventHandler =
                Sub(sender, args)
                    protocolViewer.CaretIndex = protocolViewer.Text.Length
                    protocolViewer.ScrollToEnd()
                End Sub
            AddHandler protocolViewer.TextChanged, keepScrolledToBottom

            Dim scrollOnLoad As RoutedEventHandler =
                Sub(sender, args)
                    protocolViewer.CaretIndex = protocolViewer.Text.Length
                    protocolViewer.ScrollToEnd()
                End Sub
            AddHandler protocolViewer.Loaded, scrollOnLoad

            Dim viewerChrome As New Border() With {
                .Background = ResolveBrush("SurfaceBrush", Brushes.White),
                .BorderBrush = ResolveBrush("BorderBrush", Brushes.LightGray),
                .BorderThickness = New Thickness(1),
                .CornerRadius = New CornerRadius(12),
                .Padding = New Thickness(10),
                .Child = protocolViewer
            }

            Dim headerTitle As New TextBlock() With {
                .Text = "Protocol",
                .FontSize = 16,
                .FontWeight = FontWeights.SemiBold
            }

            Dim headerCaption As New TextBlock() With {
                .Text = "Live request/response log",
                .Margin = New Thickness(0, 2, 0, 0)
            }
            headerCaption.Style = TryCast(TryFindResource("CaptionTextStyle"), Style)

            Dim headerStack As New StackPanel()
            headerStack.Children.Add(headerTitle)
            headerStack.Children.Add(headerCaption)

            Dim layout As New Grid()
            layout.RowDefinitions.Add(New RowDefinition() With {.Height = GridLength.Auto})
            layout.RowDefinitions.Add(New RowDefinition() With {.Height = New GridLength(1, GridUnitType.Star)})
            layout.RowDefinitions.Add(New RowDefinition() With {.Height = GridLength.Auto})

            Grid.SetRow(headerStack, 0)
            layout.Children.Add(headerStack)

            viewerChrome.Margin = New Thickness(0, 10, 0, 10)
            Grid.SetRow(viewerChrome, 1)
            layout.Children.Add(viewerChrome)

            Grid.SetRow(closeButton, 2)
            layout.Children.Add(closeButton)

            Dim dialog As New Window() With {
                .Owner = Me,
                .Title = "Protocol",
                .Width = 980,
                .Height = 680,
                .MinWidth = 680,
                .MinHeight = 460,
                .WindowStartupLocation = WindowStartupLocation.CenterOwner,
                .Background = ResolveBrush("AppBackgroundBrush", Brushes.White),
                .Content = New Border() With {
                    .Padding = New Thickness(14),
                    .Child = layout
                }
            }

            AddHandler closeButton.Click, Sub(sender, e) dialog.Close()

            AddHandler dialog.Closed,
                Sub(sender, e)
                    RemoveHandler protocolViewer.TextChanged, keepScrolledToBottom
                    RemoveHandler protocolViewer.Loaded, scrollOnLoad
                    BindingOperations.ClearBinding(protocolViewer, TextBox.TextProperty)
                    _protocolDialogWindow = Nothing
                End Sub

            _protocolDialogWindow = dialog
            dialog.Show()
            dialog.Activate()
        End Sub

        Private Sub OnThreadSortMenuButtonClick(sender As Object, e As RoutedEventArgs)
            OpenThreadToolbarMenu(TryCast(sender, Button), SidebarPaneHost.ThreadSortContextMenu)
        End Sub

        Private Sub OnThreadFilterMenuButtonClick(sender As Object, e As RoutedEventArgs)
            OpenThreadToolbarMenu(TryCast(sender, Button), SidebarPaneHost.ThreadFilterContextMenu)
        End Sub

        Private Sub OpenThreadToolbarMenu(button As Button, menu As ContextMenu)
            If button Is Nothing OrElse menu Is Nothing OrElse Not button.IsEnabled Then
                Return
            End If

            SyncThreadToolbarMenus()

            If menu.IsOpen Then
                menu.IsOpen = False
                Return
            End If

            menu.PlacementTarget = button
            menu.IsOpen = True
        End Sub

        Private Sub SyncThreadToolbarMenus()
            If SidebarPaneHost.MnuThreadSortNewest Is Nothing OrElse
               SidebarPaneHost.MnuThreadSortOldest Is Nothing OrElse
               SidebarPaneHost.MnuThreadSortPreviewAz Is Nothing OrElse
               SidebarPaneHost.MnuThreadSortPreviewZa Is Nothing OrElse
               SidebarPaneHost.MnuThreadFilterArchived Is Nothing OrElse
               SidebarPaneHost.MnuThreadFilterWorkingDir Is Nothing Then
                Return
            End If

            _suppressThreadToolbarMenuEvents = True
            Try
                Dim sortIndex = Math.Max(0, _viewModel.ThreadsPanel.SortIndex)
                SidebarPaneHost.MnuThreadSortNewest.IsChecked = (sortIndex = 0)
                SidebarPaneHost.MnuThreadSortOldest.IsChecked = (sortIndex = 1)
                SidebarPaneHost.MnuThreadSortPreviewAz.IsChecked = (sortIndex = 2)
                SidebarPaneHost.MnuThreadSortPreviewZa.IsChecked = (sortIndex = 3)

                SidebarPaneHost.MnuThreadFilterArchived.IsChecked = _viewModel.ThreadsPanel.ShowArchived
                SidebarPaneHost.MnuThreadFilterWorkingDir.IsChecked = _viewModel.ThreadsPanel.FilterByWorkingDir
            Finally
                _suppressThreadToolbarMenuEvents = False
            End Try
        End Sub

        Private Sub OnThreadSortMenuItemClick(sender As Object, e As RoutedEventArgs)
            Dim targetIndex As Integer = -1
            If ReferenceEquals(sender, SidebarPaneHost.MnuThreadSortNewest) Then
                targetIndex = 0
            ElseIf ReferenceEquals(sender, SidebarPaneHost.MnuThreadSortOldest) Then
                targetIndex = 1
            ElseIf ReferenceEquals(sender, SidebarPaneHost.MnuThreadSortPreviewAz) Then
                targetIndex = 2
            ElseIf ReferenceEquals(sender, SidebarPaneHost.MnuThreadSortPreviewZa) Then
                targetIndex = 3
            End If

            If targetIndex < 0 Then
                Return
            End If

            ApplyThreadSortMenuSelection(targetIndex)
        End Sub

        Private Sub OnThreadFilterMenuItemToggled(sender As Object, e As RoutedEventArgs)
            ApplyThreadFilterMenuToggle()
        End Sub

        Private Sub ApplyThreadSortMenuSelection(targetIndex As Integer)
            If _suppressThreadToolbarMenuEvents Then
                Return
            End If

            If _viewModel.ThreadsPanel.SortIndex <> targetIndex Then
                _viewModel.ThreadsPanel.SortIndex = targetIndex
            Else
                SyncThreadToolbarMenus()
            End If

            If SidebarPaneHost.ThreadSortContextMenu IsNot Nothing Then
                SidebarPaneHost.ThreadSortContextMenu.IsOpen = False
            End If
        End Sub

        Private Sub ApplyThreadFilterMenuToggle()
            If _suppressThreadToolbarMenuEvents Then
                Return
            End If

            Dim archivedChecked = SidebarPaneHost.MnuThreadFilterArchived IsNot Nothing AndAlso SidebarPaneHost.MnuThreadFilterArchived.IsChecked
            Dim workingDirChecked = SidebarPaneHost.MnuThreadFilterWorkingDir IsNot Nothing AndAlso SidebarPaneHost.MnuThreadFilterWorkingDir.IsChecked

            Dim changed As Boolean = False
            If _viewModel.ThreadsPanel.ShowArchived <> archivedChecked Then
                _viewModel.ThreadsPanel.ShowArchived = archivedChecked
                changed = True
            End If

            If _viewModel.ThreadsPanel.FilterByWorkingDir <> workingDirChecked Then
                _viewModel.ThreadsPanel.FilterByWorkingDir = workingDirChecked
                changed = True
            End If

            If changed Then
                SaveSettings()
            End If

            SyncThreadToolbarMenus()
        End Sub

        Private Sub SyncSidebarNewThreadMenu()
            If SidebarPaneHost.MnuSidebarNewThreadChooseFolder Is Nothing Then
                Return
            End If

            SidebarPaneHost.MnuSidebarNewThreadChooseFolder.IsEnabled = _viewModel.IsSidebarNewThreadEnabled
        End Sub

        Private Sub OpenWorkspaceInVsCode()
            Dim targetCwd = ResolveQuickOpenWorkspaceCwd("VS Code")
            If String.IsNullOrWhiteSpace(targetCwd) Then
                Return
            End If

            If StartVsCode(targetCwd, ".") Then
                ShowStatus($"Opened VS Code in {targetCwd}")
                Return
            End If

            ShowStatus("Could not open VS Code. Make sure `code` is installed and available on PATH.", isError:=True, displayToast:=True)
        End Sub

        Private Sub OpenWorkspaceInPowerShell()
            Dim targetCwd = ResolveQuickOpenWorkspaceCwd("PowerShell")
            If String.IsNullOrWhiteSpace(targetCwd) Then
                Return
            End If

            If StartProcessInDirectory("powershell.exe", "-NoExit", targetCwd) Then
                ShowStatus($"Opened PowerShell in {targetCwd}")
                Return
            End If

            ShowStatus("Could not open PowerShell.", isError:=True, displayToast:=True)
        End Sub

        Private Sub InitializeGitPanelUi()
            ShowGitPanelTab(_gitPanelActiveTab)
            ResetGitPanelTabContent()
        End Sub

        Private Sub ShowGitPanelTab(tabKey As String)
            _gitPanelActiveTab = If(String.IsNullOrWhiteSpace(tabKey), "changes", tabKey.Trim().ToLowerInvariant())

            If GitPaneHost.GitTabChangesView IsNot Nothing Then
                GitPaneHost.GitTabChangesView.Visibility = If(StringComparer.Ordinal.Equals(_gitPanelActiveTab, "changes"), Visibility.Visible, Visibility.Collapsed)
            End If
            If GitPaneHost.GitTabHistoryView IsNot Nothing Then
                GitPaneHost.GitTabHistoryView.Visibility = If(StringComparer.Ordinal.Equals(_gitPanelActiveTab, "history"), Visibility.Visible, Visibility.Collapsed)
            End If
            If GitPaneHost.GitTabBranchesView IsNot Nothing Then
                GitPaneHost.GitTabBranchesView.Visibility = If(StringComparer.Ordinal.Equals(_gitPanelActiveTab, "branches"), Visibility.Visible, Visibility.Collapsed)
            End If

            ApplyGitPanelTabButtonState(GitPaneHost.BtnGitTabChanges, StringComparer.Ordinal.Equals(_gitPanelActiveTab, "changes"))
            ApplyGitPanelTabButtonState(GitPaneHost.BtnGitTabHistory, StringComparer.Ordinal.Equals(_gitPanelActiveTab, "history"))
            ApplyGitPanelTabButtonState(GitPaneHost.BtnGitTabBranches, StringComparer.Ordinal.Equals(_gitPanelActiveTab, "branches"))
        End Sub

        Private Sub ApplyGitPanelTabButtonState(button As Button, isSelected As Boolean)
            If button Is Nothing Then
                Return
            End If

            button.Background = ResolveBrush(If(isSelected, "AccentSubtleBrush", "SurfaceBrush"), Brushes.Transparent)
            button.BorderBrush = ResolveBrush(If(isSelected, "AccentGlowBrush", "BorderBrush"), Brushes.Transparent)
            button.BorderThickness = New Thickness(1)
            button.Foreground = ResolveBrush(If(isSelected, "TextSecondaryBrush", "TextTertiaryBrush"), Brushes.Black)
            button.FontWeight = If(isSelected, FontWeights.SemiBold, FontWeights.Normal)
        End Sub

        Private Sub ResetGitPanelTabContent()
            _suppressGitPanelSelectionEvents = True
            Try
                If GitPaneHost.LstGitPanelChanges IsNot Nothing Then
                    GitPaneHost.LstGitPanelChanges.ItemsSource = Nothing
                    GitPaneHost.LstGitPanelChanges.SelectedItem = Nothing
                End If
                If GitPaneHost.LstGitPanelCommits IsNot Nothing Then
                    GitPaneHost.LstGitPanelCommits.ItemsSource = Nothing
                    GitPaneHost.LstGitPanelCommits.SelectedItem = Nothing
                End If
                If GitPaneHost.LstGitPanelBranches IsNot Nothing Then
                    GitPaneHost.LstGitPanelBranches.ItemsSource = Nothing
                    GitPaneHost.LstGitPanelBranches.SelectedItem = Nothing
                End If
            Finally
                _suppressGitPanelSelectionEvents = False
            End Try

            SetGitPanelDiffPreviewText("Select a changed file to preview its diff.")
            If GitPaneHost.TxtGitPanelCommitPreview IsNot Nothing Then
                GitPaneHost.TxtGitPanelCommitPreview.Text = "Select a commit to preview details."
            End If
            If GitPaneHost.TxtGitPanelBranchPreview IsNot Nothing Then
                GitPaneHost.TxtGitPanelBranchPreview.Text = "Select a branch to preview recent history."
            End If
            If GitPaneHost.LblGitPanelDiffTitle IsNot Nothing Then
                GitPaneHost.LblGitPanelDiffTitle.Text = "Diff Preview"
            End If
            If GitPaneHost.LblGitPanelDiffMeta IsNot Nothing Then
                GitPaneHost.LblGitPanelDiffMeta.Text = String.Empty
            End If
            If GitPaneHost.LblGitPanelCommitPreviewTitle IsNot Nothing Then
                GitPaneHost.LblGitPanelCommitPreviewTitle.Text = "Commit Preview"
            End If
            If GitPaneHost.LblGitPanelBranchPreviewTitle IsNot Nothing Then
                GitPaneHost.LblGitPanelBranchPreviewTitle.Text = "Branch Preview"
            End If
        End Sub

        Private Sub OnGitPanelChangesSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If _suppressGitPanelSelectionEvents Then
                Return
            End If

            Dim selected = TryCast(GitPaneHost.LstGitPanelChanges?.SelectedItem, GitChangedFileListEntry)
            If selected Is Nothing Then
                SetGitPanelDiffPreviewText("Select a changed file to preview its diff.")
                Return
            End If

            _gitPanelSelectedDiffFilePath = If(selected.FilePath, String.Empty)

            FireAndForget(LoadGitChangeDiffPreviewAsync(selected))
        End Sub

        Private Sub OnGitPanelChangesMouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
            If _suppressGitPanelSelectionEvents Then
                Return
            End If

            Dim selected = TryCast(GitPaneHost.LstGitPanelChanges?.SelectedItem, GitChangedFileListEntry)
            If selected Is Nothing Then
                Return
            End If

            OpenGitChangeInVsCode(selected)
        End Sub

        Private Sub OnGitPanelChangesListButtonClick(sender As Object, e As RoutedEventArgs)
            If _suppressGitPanelSelectionEvents OrElse e Is Nothing Then
                Return
            End If

            Dim source = TryCast(e.OriginalSource, DependencyObject)
            If source Is Nothing Then
                Return
            End If

            Dim button = FindVisualAncestor(Of Button)(source)
            If button Is Nothing OrElse Not StringComparer.Ordinal.Equals(button.Name, "BtnGitChangeOpenInline") Then
                Return
            End If

            Dim selected = TryCast(button.Tag, GitChangedFileListEntry)
            If selected Is Nothing Then
                selected = TryCast(button.DataContext, GitChangedFileListEntry)
            End If
            If selected Is Nothing Then
                Return
            End If

            e.Handled = True
            OpenGitChangeInVsCode(selected)
        End Sub

        Private Sub OnGitPanelCommitsSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If _suppressGitPanelSelectionEvents Then
                Return
            End If

            Dim selected = TryCast(GitPaneHost.LstGitPanelCommits?.SelectedItem, GitCommitListEntry)
            If selected Is Nothing Then
                If GitPaneHost.TxtGitPanelCommitPreview IsNot Nothing Then
                    GitPaneHost.TxtGitPanelCommitPreview.Text = "Select a commit to preview details."
                End If
                Return
            End If

            FireAndForget(LoadGitCommitPreviewAsync(selected))
        End Sub

        Private Sub OnGitPanelBranchesSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If _suppressGitPanelSelectionEvents Then
                Return
            End If

            Dim selected = TryCast(GitPaneHost.LstGitPanelBranches?.SelectedItem, GitBranchListEntry)
            If selected Is Nothing Then
                If GitPaneHost.TxtGitPanelBranchPreview IsNot Nothing Then
                    GitPaneHost.TxtGitPanelBranchPreview.Text = "Select a branch to preview recent history."
                End If
                Return
            End If

            FireAndForget(LoadGitBranchPreviewAsync(selected))
        End Sub

        Private Sub SetGitPanelDiffPreviewText(previewText As String)
            If GitPaneHost.LstGitPanelDiffPreviewLines Is Nothing Then
                Return
            End If

            Dim lines = BuildGitDiffPreviewLineEntries(previewText)
            GitPaneHost.LstGitPanelDiffPreviewLines.ItemsSource = Nothing
            GitPaneHost.LstGitPanelDiffPreviewLines.ItemsSource = lines

            If lines IsNot Nothing AndAlso lines.Count > 0 Then
                GitPaneHost.LstGitPanelDiffPreviewLines.ScrollIntoView(lines(0))
            End If
        End Sub

        Private Shared Function BuildGitDiffPreviewLineEntries(previewText As String) As List(Of GitDiffPreviewLineEntry)
            Dim text = If(previewText, String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            If text.Length = 0 Then
                Return New List(Of GitDiffPreviewLineEntry) From {
                    New GitDiffPreviewLineEntry() With {.Text = String.Empty, .Kind = "message"}
                }
            End If

            Dim result As New List(Of GitDiffPreviewLineEntry)()
            Dim lines = text.Split({vbLf}, StringSplitOptions.None)
            For Each line In lines
                result.Add(New GitDiffPreviewLineEntry() With {
                    .Text = line,
                    .Kind = ClassifyGitDiffPreviewLine(line)
                })
            Next

            Return result
        End Function

        Private Shared Function ClassifyGitDiffPreviewLine(line As String) As String
            Dim value = If(line, String.Empty)
            If value.Length = 0 Then
                Return "empty"
            End If

            If value.StartsWith("diff --git ", StringComparison.Ordinal) Then
                Return "fileHeader"
            End If

            If value.StartsWith("--- ", StringComparison.Ordinal) OrElse
               value.StartsWith("+++ ", StringComparison.Ordinal) Then
                Return "pathHeader"
            End If

            If value.StartsWith("@@", StringComparison.Ordinal) Then
                Return "hunk"
            End If

            If value.StartsWith("+", StringComparison.Ordinal) AndAlso Not value.StartsWith("+++", StringComparison.Ordinal) Then
                Return "added"
            End If

            If value.StartsWith("-", StringComparison.Ordinal) AndAlso Not value.StartsWith("---", StringComparison.Ordinal) Then
                Return "removed"
            End If

            If value.StartsWith("index ", StringComparison.Ordinal) OrElse
               value.StartsWith("new file mode ", StringComparison.Ordinal) OrElse
               value.StartsWith("deleted file mode ", StringComparison.Ordinal) OrElse
               value.StartsWith("old mode ", StringComparison.Ordinal) OrElse
               value.StartsWith("new mode ", StringComparison.Ordinal) OrElse
               value.StartsWith("similarity index ", StringComparison.Ordinal) OrElse
               value.StartsWith("rename from ", StringComparison.Ordinal) OrElse
               value.StartsWith("rename to ", StringComparison.Ordinal) OrElse
               value.StartsWith("copy from ", StringComparison.Ordinal) OrElse
               value.StartsWith("copy to ", StringComparison.Ordinal) OrElse
               value.StartsWith("Binary files ", StringComparison.Ordinal) Then
                Return "meta"
            End If

            If value.StartsWith("\", StringComparison.Ordinal) Then
                Return "note"
            End If

            If value.StartsWith(" ", StringComparison.Ordinal) Then
                Return "context"
            End If

            Return "message"
        End Function

        Private Async Function LoadGitChangeDiffPreviewAsync(selected As GitChangedFileListEntry) As Task
            If selected Is Nothing Then
                Return
            End If

            Dim repoRoot = ResolveCurrentGitPanelRepoRoot()
            If String.IsNullOrWhiteSpace(repoRoot) Then
                Return
            End If

            If GitPaneHost.LblGitPanelDiffTitle IsNot Nothing Then
                GitPaneHost.LblGitPanelDiffTitle.Text = $"Diff: {selected.DisplayPath}"
            End If
            If GitPaneHost.LblGitPanelDiffMeta IsNot Nothing Then
                Dim metaParts As New List(Of String)()
                If selected.AddedLineCount.HasValue AndAlso selected.AddedLineCount.Value > 0 Then
                    metaParts.Add($"+{selected.AddedLineCount.Value}")
                End If
                If selected.RemovedLineCount.HasValue AndAlso selected.RemovedLineCount.Value > 0 Then
                    metaParts.Add($"-{selected.RemovedLineCount.Value}")
                End If
                GitPaneHost.LblGitPanelDiffMeta.Text = String.Join("  ", metaParts)
            End If
            SetGitPanelDiffPreviewText("Loading diff preview...")

            Dim previewVersion = Interlocked.Increment(_gitPanelDiffPreviewLoadVersion)
            Dim previewText = Await Task.Run(Function() BuildGitFileDiffPreview(repoRoot, selected)).ConfigureAwait(True)

            If previewVersion <> _gitPanelDiffPreviewLoadVersion Then
                Return
            End If

            SetGitPanelDiffPreviewText(previewText)
        End Function

        Private Async Function LoadGitCommitPreviewAsync(selected As GitCommitListEntry) As Task
            If selected Is Nothing Then
                Return
            End If

            Dim repoRoot = ResolveCurrentGitPanelRepoRoot()
            If String.IsNullOrWhiteSpace(repoRoot) Then
                Return
            End If

            If GitPaneHost.LblGitPanelCommitPreviewTitle IsNot Nothing Then
                GitPaneHost.LblGitPanelCommitPreviewTitle.Text = $"Commit Preview: {selected.ShortSha}"
            End If
            If GitPaneHost.TxtGitPanelCommitPreview IsNot Nothing Then
                GitPaneHost.TxtGitPanelCommitPreview.Text = "Loading commit preview..."
            End If

            Dim previewVersion = Interlocked.Increment(_gitPanelCommitPreviewLoadVersion)
            Dim previewText = Await Task.Run(Function() BuildGitCommitPreview(repoRoot, selected.Sha)).ConfigureAwait(True)
            If previewVersion <> _gitPanelCommitPreviewLoadVersion Then
                Return
            End If

            If GitPaneHost.TxtGitPanelCommitPreview IsNot Nothing Then
                GitPaneHost.TxtGitPanelCommitPreview.Text = previewText
            End If
        End Function

        Private Async Function LoadGitBranchPreviewAsync(selected As GitBranchListEntry) As Task
            If selected Is Nothing Then
                Return
            End If

            Dim repoRoot = ResolveCurrentGitPanelRepoRoot()
            If String.IsNullOrWhiteSpace(repoRoot) Then
                Return
            End If

            If GitPaneHost.LblGitPanelBranchPreviewTitle IsNot Nothing Then
                GitPaneHost.LblGitPanelBranchPreviewTitle.Text = $"Branch Preview: {selected.Name}"
            End If
            If GitPaneHost.TxtGitPanelBranchPreview IsNot Nothing Then
                GitPaneHost.TxtGitPanelBranchPreview.Text = "Loading branch preview..."
            End If

            Dim previewVersion = Interlocked.Increment(_gitPanelBranchPreviewLoadVersion)
            Dim previewText = Await Task.Run(Function() BuildGitBranchPreview(repoRoot, selected.Name)).ConfigureAwait(True)
            If previewVersion <> _gitPanelBranchPreviewLoadVersion Then
                Return
            End If

            If GitPaneHost.TxtGitPanelBranchPreview IsNot Nothing Then
                GitPaneHost.TxtGitPanelBranchPreview.Text = previewText
            End If
        End Function

        Private Function ResolveCurrentGitPanelRepoRoot() As String
            If _currentGitPanelSnapshot Is Nothing Then
                Return String.Empty
            End If

            Return If(_currentGitPanelSnapshot.RepoRoot, String.Empty)
        End Function

        Private Sub OpenGitChangeInVsCode(selected As GitChangedFileListEntry)
            If selected Is Nothing Then
                Return
            End If

            Dim repoRoot = ResolveCurrentGitPanelRepoRoot()
            If String.IsNullOrWhiteSpace(repoRoot) Then
                ShowStatus("Git panel repository is not available.", isError:=True, displayToast:=True)
                Return
            End If

            Dim relativePath = If(selected.FilePath, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(relativePath) Then
                ShowStatus("No file path available for this entry.", isError:=True, displayToast:=True)
                Return
            End If

            Dim fullPath = Path.Combine(repoRoot, relativePath.Replace("/"c, Path.DirectorySeparatorChar))
            If Not File.Exists(fullPath) Then
                ShowStatus($"File not found (possibly deleted): {relativePath}", isError:=True, displayToast:=True)
                Return
            End If

            If StartVsCode(repoRoot, "-g " & QuoteProcessArgument(fullPath)) Then
                ShowStatus($"Opened {relativePath} in VS Code")
                Return
            End If

            ShowStatus("Could not open file in VS Code. Make sure `code` is installed and available on PATH.",
                       isError:=True,
                       displayToast:=True)
        End Sub

        Private Sub ToggleGitPanel()
            If GitPaneHost.GitInspectorPanel Is Nothing Then
                Return
            End If

            If GitPaneHost.GitInspectorPanel.Visibility = Visibility.Visible Then
                CloseGitPanel()
                Return
            End If

            ShowGitPanelDock()
            GitPaneHost.GitInspectorPanel.Visibility = Visibility.Visible
            FireAndForget(RefreshGitPanelAsync())
        End Sub

        Private Sub ShowGitPanelDock()
            Const minGitPaneWidth As Double = 280.0R
            Const preferredGitPaneWidth As Double = 560.0R
            Const maxGitPaneWidth As Double = 760.0R

            Dim targetWidth = _gitPanelDockWidth
            If Double.IsNaN(targetWidth) OrElse Double.IsInfinity(targetWidth) OrElse targetWidth < minGitPaneWidth Then
                targetWidth = preferredGitPaneWidth
            End If

            ' Always reopen at a comfortable width when possible, even if the user last closed it very narrow.
            targetWidth = Math.Max(targetWidth, preferredGitPaneWidth)

            Dim maxFitWidth = Double.PositiveInfinity
            If WorkspacePaneHost IsNot Nothing AndAlso WorkspacePaneHost.ActualWidth > 0 Then
                maxFitWidth = Math.Max(minGitPaneWidth, WorkspacePaneHost.ActualWidth - WorkspaceResponsiveMinWidthWhenGitPaneOpen)
            End If

            If Not Double.IsInfinity(maxFitWidth) Then
                targetWidth = Math.Min(targetWidth, maxFitWidth)
            End If

            If MainSurfaceGrid IsNot Nothing AndAlso MainSurfaceGrid.ActualWidth > 0 Then
                Dim maxHalfWidth = Math.Floor(MainSurfaceGrid.ActualWidth / 2.0R)
                If maxHalfWidth > 0 Then
                    targetWidth = Math.Min(targetWidth, maxHalfWidth)
                End If
            End If

            targetWidth = Math.Max(minGitPaneWidth, Math.Min(targetWidth, maxGitPaneWidth))

            If RightGitPaneShell IsNot Nothing Then
                RightGitPaneShell.Visibility = Visibility.Visible
            End If

            If RightGitPaneColumn IsNot Nothing Then
                RightGitPaneColumn.MinWidth = minGitPaneWidth
                RightGitPaneColumn.Width = New GridLength(targetWidth, GridUnitType.Pixel)
            End If

            If RightGitPaneSplitterColumn IsNot Nothing Then
                RightGitPaneSplitterColumn.Width = New GridLength(8, GridUnitType.Pixel)
            End If

            If RightGitPaneSplitter IsNot Nothing Then
                RightGitPaneSplitter.Visibility = Visibility.Visible
            End If

            UpdateMainPaneResizeBounds()
        End Sub

        Private Shared Function StartVsCode(workingDirectory As String, arguments As String) As Boolean
            Dim normalizedArgs = If(arguments, String.Empty)

            ' Prefer direct executables first to avoid a visible cmd window.
            Dim directCandidates As New List(Of String) From {
                "code.exe",
                "code"
            }

            Dim userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            Dim programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
            Dim programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)

            directCandidates.Add(Path.Combine(localAppData, "Programs", "Microsoft VS Code", "Code.exe"))
            directCandidates.Add(Path.Combine(userProfile, "AppData", "Local", "Programs", "Microsoft VS Code", "Code.exe"))
            directCandidates.Add(Path.Combine(programFiles, "Microsoft VS Code", "Code.exe"))
            If Not String.IsNullOrWhiteSpace(programFilesX86) Then
                directCandidates.Add(Path.Combine(programFilesX86, "Microsoft VS Code", "Code.exe"))
            End If

            For Each candidate In directCandidates
                If String.IsNullOrWhiteSpace(candidate) Then
                    Continue For
                End If

                If candidate.Contains(Path.DirectorySeparatorChar) OrElse candidate.Contains(Path.AltDirectorySeparatorChar) Then
                    If Not File.Exists(candidate) Then
                        Continue For
                    End If
                End If

                If StartProcessInDirectory(candidate, normalizedArgs, workingDirectory, createNoWindow:=True) Then
                    Return True
                End If
            Next

            ' Fallback for PATH setups that expose only code.cmd; keep the shell window hidden.
            Return StartProcessInDirectory("cmd.exe", "/c code " & normalizedArgs, workingDirectory, createNoWindow:=True)
        End Function

        Private Sub CloseGitPanel()
            If GitPaneHost.GitInspectorPanel Is Nothing Then
                Return
            End If

            Dim actualWidth = GitPaneHost.GitInspectorPanel.ActualWidth
            If Not Double.IsNaN(actualWidth) AndAlso Not Double.IsInfinity(actualWidth) AndAlso actualWidth >= 280 Then
                _gitPanelDockWidth = actualWidth
            End If

            GitPaneHost.GitInspectorPanel.Visibility = Visibility.Collapsed
            If RightGitPaneSplitter IsNot Nothing Then
                RightGitPaneSplitter.Visibility = Visibility.Collapsed
            End If
            If RightGitPaneSplitterColumn IsNot Nothing Then
                RightGitPaneSplitterColumn.Width = New GridLength(0, GridUnitType.Pixel)
            End If
            If RightGitPaneColumn IsNot Nothing Then
                RightGitPaneColumn.MinWidth = 0
                RightGitPaneColumn.Width = New GridLength(0, GridUnitType.Pixel)
            End If
            If RightGitPaneShell IsNot Nothing Then
                RightGitPaneShell.Visibility = Visibility.Collapsed
            End If
            Interlocked.Increment(_gitPanelLoadVersion)
            Interlocked.Increment(_gitPanelDiffPreviewLoadVersion)
            Interlocked.Increment(_gitPanelCommitPreviewLoadVersion)
            Interlocked.Increment(_gitPanelBranchPreviewLoadVersion)
        End Sub

        Private Async Function RefreshGitPanelAsync() As Task
            If GitPaneHost.GitInspectorPanel Is Nothing OrElse
               GitPaneHost.GitInspectorPanel.Visibility <> Visibility.Visible Then
                Return
            End If

            Dim targetCwd = ResolveQuickOpenWorkspaceCwd("Git panel")
            If String.IsNullOrWhiteSpace(targetCwd) Then
                SetGitPanelErrorState("No workspace folder available.")
                Return
            End If

            Dim loadVersion = Interlocked.Increment(_gitPanelLoadVersion)
            SetGitPanelLoadingState(True, $"Loading repository in {targetCwd}...")

            Dim snapshot = Await Task.Run(Function() BuildGitPanelSnapshot(targetCwd)).ConfigureAwait(True)

            If loadVersion <> _gitPanelLoadVersion Then
                Return
            End If

            If GitPaneHost.GitInspectorPanel.Visibility <> Visibility.Visible Then
                Return
            End If

            ApplyGitPanelSnapshot(snapshot)
        End Function

        Private Sub SetGitPanelLoadingState(isLoading As Boolean, statusText As String)
            If GitPaneHost.GitPanelLoadingOverlay IsNot Nothing Then
                GitPaneHost.GitPanelLoadingOverlay.Visibility = If(isLoading, Visibility.Visible, Visibility.Collapsed)
            End If

            If GitPaneHost.LblGitPanelLoading IsNot Nothing Then
                GitPaneHost.LblGitPanelLoading.Text = If(statusText, String.Empty)
            End If

            If GitPaneHost.LblGitPanelState IsNot Nothing Then
                GitPaneHost.LblGitPanelState.Text = If(statusText, String.Empty)
            End If

            If GitPaneHost.BtnGitPanelRefresh IsNot Nothing Then
                GitPaneHost.BtnGitPanelRefresh.IsEnabled = Not isLoading
            End If
        End Sub

        Private Sub SetGitPanelErrorState(message As String)
            _currentGitPanelSnapshot = Nothing
            SetGitPanelLoadingState(False, If(message, "Git panel unavailable."))
            If GitPaneHost.LblGitPanelRepoName IsNot Nothing Then
                GitPaneHost.LblGitPanelRepoName.Text = "Repository unavailable"
            End If
            If GitPaneHost.LblGitPanelRepoPath IsNot Nothing Then
                GitPaneHost.LblGitPanelRepoPath.Text = String.Empty
            End If
            If GitPaneHost.LblGitPanelBranch IsNot Nothing Then
                GitPaneHost.LblGitPanelBranch.Text = "branch: -"
            End If
            If GitPaneHost.LblGitPanelStatusSummary IsNot Nothing Then
                GitPaneHost.LblGitPanelStatusSummary.Text = "status: unavailable"
            End If
            SetGitPanelLineSummary(Nothing, Nothing)
            ResetGitPanelTabContent()
            SetGitPanelDiffPreviewText(If(message, "No git data."))
            If GitPaneHost.TxtGitPanelCommitPreview IsNot Nothing Then
                GitPaneHost.TxtGitPanelCommitPreview.Text = "No commit history available."
            End If
            If GitPaneHost.TxtGitPanelBranchPreview IsNot Nothing Then
                GitPaneHost.TxtGitPanelBranchPreview.Text = "No branch data available."
            End If
        End Sub

        Private Sub ApplyGitPanelSnapshot(snapshot As GitPanelSnapshot)
            If snapshot Is Nothing Then
                SetGitPanelErrorState("No git data.")
                Return
            End If

            If Not String.IsNullOrWhiteSpace(snapshot.ErrorMessage) Then
                SetGitPanelErrorState(snapshot.ErrorMessage)
                Return
            End If

            _currentGitPanelSnapshot = snapshot
            SetGitPanelLoadingState(False, $"Updated {snapshot.LoadedAtLocal:HH:mm:ss}")

            If GitPaneHost.LblGitPanelRepoName IsNot Nothing Then
                GitPaneHost.LblGitPanelRepoName.Text = If(snapshot.RepoName, "Repository")
            End If

            If GitPaneHost.LblGitPanelRepoPath IsNot Nothing Then
                GitPaneHost.LblGitPanelRepoPath.Text = If(snapshot.RepoRoot, snapshot.WorkingDirectory)
                GitPaneHost.LblGitPanelRepoPath.ToolTip = If(snapshot.RepoRoot, snapshot.WorkingDirectory)
            End If

            If GitPaneHost.LblGitPanelBranch IsNot Nothing Then
                Dim branchLabel = If(String.IsNullOrWhiteSpace(snapshot.BranchName), "-", snapshot.BranchName)
                GitPaneHost.LblGitPanelBranch.Text = $"branch: {branchLabel}"
            End If

            If GitPaneHost.LblGitPanelStatusSummary IsNot Nothing Then
                Dim statusLabel = If(String.IsNullOrWhiteSpace(snapshot.StatusSummary), "unknown", snapshot.StatusSummary)
                GitPaneHost.LblGitPanelStatusSummary.Text = $"status: {statusLabel}"
            End If

            SetGitPanelLineSummary(snapshot.AddedLineCount, snapshot.RemovedLineCount)
            PopulateGitPanelLists(snapshot)
        End Sub

        Private Sub PopulateGitPanelLists(snapshot As GitPanelSnapshot)
            Dim desiredChangeIndex = -1
            If snapshot IsNot Nothing AndAlso snapshot.ChangedFiles IsNot Nothing AndAlso snapshot.ChangedFiles.Count > 0 Then
                If Not String.IsNullOrWhiteSpace(_gitPanelSelectedDiffFilePath) Then
                    For i = 0 To snapshot.ChangedFiles.Count - 1
                        Dim entry = snapshot.ChangedFiles(i)
                        If entry IsNot Nothing AndAlso
                           StringComparer.OrdinalIgnoreCase.Equals(If(entry.FilePath, String.Empty), _gitPanelSelectedDiffFilePath) Then
                            desiredChangeIndex = i
                            Exit For
                        End If
                    Next
                End If

                If desiredChangeIndex < 0 Then
                    desiredChangeIndex = 0
                End If
            End If

            _suppressGitPanelSelectionEvents = True
            Try
                If GitPaneHost.LstGitPanelChanges IsNot Nothing Then
                    GitPaneHost.LstGitPanelChanges.ItemsSource = Nothing
                    GitPaneHost.LstGitPanelChanges.ItemsSource = snapshot.ChangedFiles
                    GitPaneHost.LstGitPanelChanges.SelectedIndex = desiredChangeIndex
                End If

                If GitPaneHost.LstGitPanelCommits IsNot Nothing Then
                    GitPaneHost.LstGitPanelCommits.ItemsSource = Nothing
                    GitPaneHost.LstGitPanelCommits.ItemsSource = snapshot.Commits
                    GitPaneHost.LstGitPanelCommits.SelectedIndex = If(snapshot.Commits IsNot Nothing AndAlso snapshot.Commits.Count > 0, 0, -1)
                End If

                If GitPaneHost.LstGitPanelBranches IsNot Nothing Then
                    GitPaneHost.LstGitPanelBranches.ItemsSource = Nothing
                    GitPaneHost.LstGitPanelBranches.ItemsSource = snapshot.Branches
                    Dim selectedBranchIndex = -1
                    If snapshot.Branches IsNot Nothing Then
                        For i = 0 To snapshot.Branches.Count - 1
                            If snapshot.Branches(i) IsNot Nothing AndAlso snapshot.Branches(i).IsCurrent Then
                                selectedBranchIndex = i
                                Exit For
                            End If
                        Next
                    End If
                    If selectedBranchIndex < 0 AndAlso snapshot.Branches IsNot Nothing AndAlso snapshot.Branches.Count > 0 Then
                        selectedBranchIndex = 0
                    End If
                    GitPaneHost.LstGitPanelBranches.SelectedIndex = selectedBranchIndex
                End If
            Finally
                _suppressGitPanelSelectionEvents = False
            End Try

            If snapshot.ChangedFiles Is Nothing OrElse snapshot.ChangedFiles.Count = 0 Then
                SetGitPanelDiffPreviewText(If(snapshot.ChangesText, "Working tree clean."))
                If GitPaneHost.LblGitPanelDiffTitle IsNot Nothing Then
                    GitPaneHost.LblGitPanelDiffTitle.Text = "Diff Preview"
                End If
                If GitPaneHost.LblGitPanelDiffMeta IsNot Nothing Then
                    GitPaneHost.LblGitPanelDiffMeta.Text = String.Empty
                End If
            Else
                Dim selectedChange = TryCast(GitPaneHost.LstGitPanelChanges?.SelectedItem, GitChangedFileListEntry)
                If selectedChange Is Nothing AndAlso GitPaneHost.LstGitPanelChanges IsNot Nothing Then
                    GitPaneHost.LstGitPanelChanges.SelectedIndex = If(desiredChangeIndex >= 0, desiredChangeIndex, 0)
                    selectedChange = TryCast(GitPaneHost.LstGitPanelChanges.SelectedItem, GitChangedFileListEntry)
                End If

                If selectedChange Is Nothing Then
                    selectedChange = snapshot.ChangedFiles(Math.Max(0, Math.Min(snapshot.ChangedFiles.Count - 1, desiredChangeIndex)))
                End If

                _gitPanelSelectedDiffFilePath = If(selectedChange?.FilePath, String.Empty)
                FireAndForget(LoadGitChangeDiffPreviewAsync(selectedChange))
            End If

            If snapshot.Commits Is Nothing OrElse snapshot.Commits.Count = 0 Then
                If GitPaneHost.TxtGitPanelCommitPreview IsNot Nothing Then
                    GitPaneHost.TxtGitPanelCommitPreview.Text = If(snapshot.CommitsText, "No commits found.")
                End If
            Else
                FireAndForget(LoadGitCommitPreviewAsync(snapshot.Commits(0)))
            End If

            If snapshot.Branches Is Nothing OrElse snapshot.Branches.Count = 0 Then
                If GitPaneHost.TxtGitPanelBranchPreview IsNot Nothing Then
                    GitPaneHost.TxtGitPanelBranchPreview.Text = If(snapshot.BranchesText, "No branch data available.")
                End If
            Else
                Dim branchToPreview As GitBranchListEntry = Nothing
                For Each branchEntry In snapshot.Branches
                    If branchEntry IsNot Nothing AndAlso branchEntry.IsCurrent Then
                        branchToPreview = branchEntry
                        Exit For
                    End If
                Next

                If branchToPreview Is Nothing AndAlso snapshot.Branches.Count > 0 Then
                    branchToPreview = snapshot.Branches(0)
                End If

                If branchToPreview IsNot Nothing Then
                    FireAndForget(LoadGitBranchPreviewAsync(branchToPreview))
                End If
            End If
        End Sub

        Private Sub SetGitPanelLineSummary(addedLineCount As Integer?, removedLineCount As Integer?)
            If GitPaneHost.LblGitPanelAddedSummary IsNot Nothing Then
                If addedLineCount.HasValue AndAlso addedLineCount.Value > 0 Then
                    GitPaneHost.LblGitPanelAddedSummary.Text = $"+{addedLineCount.Value} lines"
                    GitPaneHost.LblGitPanelAddedSummary.Visibility = Visibility.Visible
                Else
                    GitPaneHost.LblGitPanelAddedSummary.Text = String.Empty
                    GitPaneHost.LblGitPanelAddedSummary.Visibility = Visibility.Collapsed
                End If
            End If

            If GitPaneHost.LblGitPanelRemovedSummary IsNot Nothing Then
                If removedLineCount.HasValue AndAlso removedLineCount.Value > 0 Then
                    GitPaneHost.LblGitPanelRemovedSummary.Text = $"-{removedLineCount.Value} lines"
                    GitPaneHost.LblGitPanelRemovedSummary.Visibility = Visibility.Visible
                Else
                    GitPaneHost.LblGitPanelRemovedSummary.Text = String.Empty
                    GitPaneHost.LblGitPanelRemovedSummary.Visibility = Visibility.Collapsed
                End If
            End If
        End Sub

        Private Shared Function BuildGitPanelSnapshot(targetCwd As String) As GitPanelSnapshot
            Dim snapshot As New GitPanelSnapshot() With {
                .WorkingDirectory = If(targetCwd, String.Empty),
                .LoadedAtLocal = Date.Now
            }

            Dim repoRootResult = RunProcessCapture("git", "rev-parse --show-toplevel", targetCwd)
            If repoRootResult.ExitCode <> 0 Then
                snapshot.ErrorMessage = If(String.IsNullOrWhiteSpace(repoRootResult.ErrorText),
                                           "Selected folder is not a git repository.",
                                           repoRootResult.ErrorText.Trim())
                Return snapshot
            End If

            snapshot.RepoRoot = FirstNonEmptyLine(repoRootResult.OutputText)
            If String.IsNullOrWhiteSpace(snapshot.RepoRoot) Then
                snapshot.ErrorMessage = "Could not determine repository root."
                Return snapshot
            End If

            snapshot.RepoName = Path.GetFileName(snapshot.RepoRoot.TrimEnd("\"c, "/"c))
            If String.IsNullOrWhiteSpace(snapshot.RepoName) Then
                snapshot.RepoName = snapshot.RepoRoot
            End If

            Dim branchResult = RunProcessCapture("git", "rev-parse --abbrev-ref HEAD", snapshot.RepoRoot)
            If branchResult.ExitCode = 0 Then
                snapshot.BranchName = FirstNonEmptyLine(branchResult.OutputText)
            End If

            Dim statusResult = RunProcessCapture("git", "status --short --branch", snapshot.RepoRoot)
            If statusResult.ExitCode = 0 Then
                Dim parsedStatus = ParseGitStatus(statusResult.OutputText)
                If String.IsNullOrWhiteSpace(snapshot.BranchName) Then
                    snapshot.BranchName = parsedStatus.BranchName
                End If
                snapshot.StatusSummary = parsedStatus.StatusSummary
                snapshot.ChangesText = parsedStatus.ChangesText
                snapshot.ChangedFiles = parsedStatus.ChangedFiles
                RemoveGitDirectoryEntries(snapshot.ChangedFiles, snapshot.RepoRoot)
                AssignGitFileIcons(snapshot.ChangedFiles, snapshot.RepoRoot)
            Else
                snapshot.StatusSummary = "unavailable"
                snapshot.ChangesText = NormalizeProcessError(statusResult)
            End If

            Dim numstatResult = RunProcessCapture("git", "diff --numstat HEAD", snapshot.RepoRoot)
            If numstatResult.ExitCode = 0 Then
                Dim diffTotals = CountGitNumstatTotals(numstatResult.OutputText)
                snapshot.AddedLineCount = diffTotals.AddedLineCount
                snapshot.RemovedLineCount = diffTotals.RemovedLineCount
                ApplyGitNumstatToChangedFiles(snapshot.ChangedFiles, numstatResult.OutputText)
            Else
                Dim fallbackNumstatResult = RunProcessCapture("git", "diff --numstat", snapshot.RepoRoot)
                If fallbackNumstatResult.ExitCode = 0 Then
                    Dim diffTotals = CountGitNumstatTotals(fallbackNumstatResult.OutputText)
                    snapshot.AddedLineCount = diffTotals.AddedLineCount
                    snapshot.RemovedLineCount = diffTotals.RemovedLineCount
                    ApplyGitNumstatToChangedFiles(snapshot.ChangedFiles, fallbackNumstatResult.OutputText)
                End If
            End If

            Dim untrackedTotals = ApplyUntrackedFileLineCounts(snapshot.ChangedFiles, snapshot.RepoRoot)
            If untrackedTotals.AddedLineCount.HasValue Then
                snapshot.AddedLineCount = (If(snapshot.AddedLineCount, 0) + untrackedTotals.AddedLineCount.Value)
            End If
            If untrackedTotals.RemovedLineCount.HasValue Then
                snapshot.RemovedLineCount = (If(snapshot.RemovedLineCount, 0) + untrackedTotals.RemovedLineCount.Value)
            End If

            Dim logResult = RunProcessCapture("git", "--no-pager log --pretty=format:%H%x1f%h%x1f%cr%x1f%s%x1f%D -n 30", snapshot.RepoRoot)
            If logResult.ExitCode = 0 Then
                snapshot.CommitsText = NormalizePanelMultiline(logResult.OutputText, "No commits found.")
                snapshot.Commits = ParseGitCommits(logResult.OutputText)
            Else
                snapshot.CommitsText = NormalizeProcessError(logResult)
            End If

            Dim branchesResult = RunProcessCapture("git", "--no-pager for-each-ref --format=%(HEAD)%x1f%(refname:short)%x1f%(objectname:short)%x1f%(committerdate:relative)%x1f%(contents:subject) refs/heads refs/remotes", snapshot.RepoRoot)
            If branchesResult.ExitCode = 0 Then
                snapshot.BranchesText = NormalizePanelMultiline(branchesResult.OutputText, "No branches found.")
                snapshot.Branches = ParseGitBranches(branchesResult.OutputText)
            Else
                snapshot.BranchesText = NormalizeProcessError(branchesResult)
            End If

            If String.IsNullOrWhiteSpace(snapshot.StatusSummary) Then
                snapshot.StatusSummary = "unknown"
            End If
            If String.IsNullOrWhiteSpace(snapshot.ChangesText) Then
                snapshot.ChangesText = "Working tree clean."
            End If
            If String.IsNullOrWhiteSpace(snapshot.CommitsText) Then
                snapshot.CommitsText = "No commits found."
            End If
            If String.IsNullOrWhiteSpace(snapshot.BranchesText) Then
                snapshot.BranchesText = "No branches found."
            End If

            Return snapshot
        End Function

        Private Shared Function RunProcessCapture(fileName As String, arguments As String, workingDirectory As String) As ProcessCaptureResult
            Try
                Using process As New Process()
                    process.StartInfo = New ProcessStartInfo(fileName) With {
                        .Arguments = If(arguments, String.Empty),
                        .WorkingDirectory = If(workingDirectory, String.Empty),
                        .UseShellExecute = False,
                        .RedirectStandardOutput = True,
                        .RedirectStandardError = True,
                        .CreateNoWindow = True
                    }

                    process.Start()
                    Dim output = process.StandardOutput.ReadToEnd()
                    Dim [error] = process.StandardError.ReadToEnd()
                    process.WaitForExit()

                    Return New ProcessCaptureResult() With {
                        .ExitCode = process.ExitCode,
                        .OutputText = If(output, String.Empty),
                        .ErrorText = If([error], String.Empty)
                    }
                End Using
            Catch ex As Exception
                Return New ProcessCaptureResult() With {
                    .ExitCode = -1,
                    .OutputText = String.Empty,
                    .ErrorText = ex.Message
                }
            End Try
        End Function

        Private Structure GitStatusParseResult
            Public Property BranchName As String
            Public Property StatusSummary As String
            Public Property ChangesText As String
            Public Property ChangedFiles As List(Of GitChangedFileListEntry)
        End Structure

        Private Structure GitNumstatTotals
            Public Property AddedLineCount As Integer?
            Public Property RemovedLineCount As Integer?
        End Structure

        Private Shared Function ParseGitStatus(outputText As String) As GitStatusParseResult
            Dim result As New GitStatusParseResult() With {
                .ChangedFiles = New List(Of GitChangedFileListEntry)()
            }
            Dim normalized = If(outputText, String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim lines = normalized.Split({vbLf}, StringSplitOptions.None)
            Dim changesBuilder As New StringBuilder()
            Dim stagedCount = 0
            Dim unstagedCount = 0
            Dim untrackedCount = 0
            Dim conflictCount = 0

            For Each rawLine In lines
                Dim line = If(rawLine, String.Empty).TrimEnd()
                If String.IsNullOrWhiteSpace(line) Then
                    Continue For
                End If

                If line.StartsWith("## ", StringComparison.Ordinal) Then
                    result.BranchName = ParseBranchNameFromStatusHeader(line)
                    Continue For
                End If

                If changesBuilder.Length > 0 Then
                    changesBuilder.AppendLine()
                End If
                changesBuilder.Append(line)

                Dim changeEntry = ParseGitStatusChangeEntry(line)
                If changeEntry IsNot Nothing Then
                    result.ChangedFiles.Add(changeEntry)
                End If

                Dim code = If(line.Length >= 2, line.Substring(0, 2), line)
                If StringComparer.Ordinal.Equals(code, "??") Then
                    untrackedCount += 1
                    Continue For
                End If

                If code.Contains("U"c) Then
                    conflictCount += 1
                End If
                If code.Length >= 1 AndAlso code(0) <> " "c Then
                    stagedCount += 1
                End If
                If code.Length >= 2 AndAlso code(1) <> " "c Then
                    unstagedCount += 1
                End If
            Next

            If changesBuilder.Length = 0 Then
                result.ChangesText = "Working tree clean."
            Else
                result.ChangesText = changesBuilder.ToString()
            End If

            Dim parts As New List(Of String)()
            If conflictCount > 0 Then
                parts.Add($"{conflictCount} conflict")
            End If
            If stagedCount > 0 Then
                parts.Add($"{stagedCount} staged")
            End If
            If unstagedCount > 0 Then
                parts.Add($"{unstagedCount} changed")
            End If
            If untrackedCount > 0 Then
                parts.Add($"{untrackedCount} untracked")
            End If

            result.StatusSummary = If(parts.Count = 0, "clean", String.Join(" • ", parts))
            Return result
        End Function

        Private Shared Function ParseGitStatusChangeEntry(statusLine As String) As GitChangedFileListEntry
            Dim line = If(statusLine, String.Empty)
            If line.StartsWith("## ", StringComparison.Ordinal) OrElse line.Length < 3 Then
                Return Nothing
            End If

            Dim rawCode = line.Substring(0, Math.Min(2, line.Length))
            Dim pathText = If(line.Length > 3, line.Substring(3).Trim(), String.Empty)
            If String.IsNullOrWhiteSpace(pathText) Then
                Return Nothing
            End If

            Dim displayPath = pathText
            Dim canonicalPath = pathText
            Dim renameArrowIndex = pathText.IndexOf(" -> ", StringComparison.Ordinal)
            If renameArrowIndex >= 0 Then
                canonicalPath = pathText.Substring(renameArrowIndex + 4).Trim()
            End If

            Dim code = rawCode.Trim()
            If String.IsNullOrWhiteSpace(code) Then
                code = rawCode
            End If

            Return New GitChangedFileListEntry() With {
                .StatusCode = If(code, String.Empty),
                .FilePath = canonicalPath,
                .DisplayPath = displayPath,
                .IsUntracked = StringComparer.Ordinal.Equals(rawCode, "??")
            }
        End Function

        Private Shared Sub RemoveGitDirectoryEntries(changedFiles As IList(Of GitChangedFileListEntry), repoRoot As String)
            If changedFiles Is Nothing OrElse changedFiles.Count = 0 Then
                Return
            End If

            For i = changedFiles.Count - 1 To 0 Step -1
                Dim item = changedFiles(i)
                If item Is Nothing OrElse String.IsNullOrWhiteSpace(item.FilePath) Then
                    Continue For
                End If

                Dim normalizedRelativePath = item.FilePath.Trim().Replace("/"c, Path.DirectorySeparatorChar)
                Dim isDirectoryHint = item.FilePath.EndsWith("/", StringComparison.Ordinal) OrElse
                                      item.FilePath.EndsWith("\", StringComparison.Ordinal)

                Dim isDirectoryOnDisk = False
                If Not String.IsNullOrWhiteSpace(repoRoot) Then
                    Dim fullPath = Path.Combine(repoRoot, normalizedRelativePath.TrimEnd(Path.DirectorySeparatorChar))
                    If Not String.IsNullOrWhiteSpace(fullPath) Then
                        isDirectoryOnDisk = Directory.Exists(fullPath)
                    End If
                End If

                If isDirectoryHint OrElse isDirectoryOnDisk Then
                    changedFiles.RemoveAt(i)
                End If
            Next
        End Sub

        Private Shared Sub AssignGitFileIcons(changedFiles As IList(Of GitChangedFileListEntry), repoRoot As String)
            If changedFiles Is Nothing OrElse changedFiles.Count = 0 Then
                Return
            End If

            For Each item In changedFiles
                If item Is Nothing Then
                    Continue For
                End If

                item.FileIconSource = GetGitFileIconSource(item, repoRoot)
            Next
        End Sub

        Private Shared Function GetGitFileIconSource(item As GitChangedFileListEntry, repoRoot As String) As ImageSource
            If item Is Nothing Then
                Return Nothing
            End If

            Dim relativePath = If(item.FilePath, String.Empty).Trim()
            Dim isDirectory = relativePath.EndsWith("/", StringComparison.Ordinal) OrElse
                              relativePath.EndsWith("\", StringComparison.Ordinal)

            Dim extension = String.Empty
            If Not isDirectory Then
                extension = Path.GetExtension(relativePath)
            End If

            Dim cacheKey As String
            If isDirectory Then
                cacheKey = "dir"
            ElseIf Not String.IsNullOrWhiteSpace(extension) Then
                cacheKey = "ext:" & extension.Trim().ToLowerInvariant()
            Else
                cacheKey = "file"
            End If

            SyncLock _gitFileIconCacheLock
                Dim cached As ImageSource = Nothing
                If _gitFileIconCache.TryGetValue(cacheKey, cached) Then
                    Return cached
                End If
            End SyncLock

            Dim iconSource = CreateShellAssociatedIconSource(relativePath, repoRoot, isDirectory)

            SyncLock _gitFileIconCacheLock
                If Not _gitFileIconCache.ContainsKey(cacheKey) Then
                    _gitFileIconCache(cacheKey) = iconSource
                End If
                Return _gitFileIconCache(cacheKey)
            End SyncLock
        End Function

        Private Shared Function CreateShellAssociatedIconSource(relativePath As String, repoRoot As String, isDirectory As Boolean) As ImageSource
            Try
                Dim pathHint As String
                If isDirectory Then
                    pathHint = "folder"
                Else
                    Dim ext = Path.GetExtension(If(relativePath, String.Empty))
                    If String.IsNullOrWhiteSpace(ext) Then
                        pathHint = "file.bin"
                    Else
                        pathHint = "file" & ext.Trim()
                    End If
                End If

                Dim fullPath = String.Empty
                If Not String.IsNullOrWhiteSpace(repoRoot) AndAlso Not String.IsNullOrWhiteSpace(relativePath) Then
                    fullPath = Path.Combine(repoRoot, relativePath.Replace("/"c, Path.DirectorySeparatorChar).TrimEnd(Path.DirectorySeparatorChar))
                End If

                Dim queryPath = If(Not String.IsNullOrWhiteSpace(fullPath), fullPath, pathHint)
                Dim fileAttrs = If(isDirectory, FILE_ATTRIBUTE_DIRECTORY, FILE_ATTRIBUTE_NORMAL)
                Dim flags = SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_USEFILEATTRIBUTES

                Dim info As New SHFILEINFO()
                Dim result = SHGetFileInfo(queryPath, fileAttrs, info, CUInt(Marshal.SizeOf(GetType(SHFILEINFO))), flags)
                If result = IntPtr.Zero OrElse info.hIcon = IntPtr.Zero Then
                    Return Nothing
                End If

                Try
                    Dim iconImage = Imaging.CreateBitmapSourceFromHIcon(info.hIcon, Int32Rect.Empty, BitmapSizeOptions.FromWidthAndHeight(16, 16))
                    If iconImage IsNot Nothing AndAlso iconImage.CanFreeze Then
                        iconImage.Freeze()
                    End If
                    Return iconImage
                Finally
                    DestroyIcon(info.hIcon)
                End Try
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function ParseBranchNameFromStatusHeader(headerLine As String) As String
            Dim header = If(headerLine, String.Empty).Trim()
            If header.StartsWith("## ", StringComparison.Ordinal) Then
                header = header.Substring(3)
            End If

            Dim ellipsisIndex = header.IndexOf("...", StringComparison.Ordinal)
            If ellipsisIndex >= 0 Then
                header = header.Substring(0, ellipsisIndex)
            End If

            Dim spaceIndex = header.IndexOf(" "c)
            If spaceIndex >= 0 Then
                header = header.Substring(0, spaceIndex)
            End If

            Return header.Trim()
        End Function

        Private Shared Sub ApplyGitNumstatToChangedFiles(changedFiles As IList(Of GitChangedFileListEntry), numstatOutput As String)
            If changedFiles Is Nothing OrElse changedFiles.Count = 0 Then
                Return
            End If

            Dim perPath = ParseGitNumstatPerPath(numstatOutput)
            If perPath.Count = 0 Then
                Return
            End If

            For Each item In changedFiles
                If item Is Nothing OrElse String.IsNullOrWhiteSpace(item.FilePath) Then
                    Continue For
                End If

                Dim stats As GitNumstatTotals = Nothing
                If perPath.TryGetValue(item.FilePath, stats) Then
                    item.AddedLineCount = stats.AddedLineCount
                    item.RemovedLineCount = stats.RemovedLineCount
                End If
            Next
        End Sub

        Private Shared Function ApplyUntrackedFileLineCounts(changedFiles As IList(Of GitChangedFileListEntry), repoRoot As String) As GitNumstatTotals
            If changedFiles Is Nothing OrElse changedFiles.Count = 0 OrElse String.IsNullOrWhiteSpace(repoRoot) Then
                Return New GitNumstatTotals()
            End If

            Dim addedTotal = 0
            Dim removedTotal = 0

            For Each item In changedFiles
                If item Is Nothing OrElse Not item.IsUntracked Then
                    Continue For
                End If

                If item.AddedLineCount.HasValue OrElse item.RemovedLineCount.HasValue Then
                    If item.AddedLineCount.HasValue Then
                        addedTotal += Math.Max(0, item.AddedLineCount.Value)
                    End If
                    If item.RemovedLineCount.HasValue Then
                        removedTotal += Math.Max(0, item.RemovedLineCount.Value)
                    End If
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(item.FilePath) Then
                    Continue For
                End If

                Dim fullPath = Path.Combine(repoRoot, item.FilePath.Replace("/"c, Path.DirectorySeparatorChar))
                Dim lineCount = CountFileLinesOrNothing(fullPath)
                If lineCount.HasValue AndAlso lineCount.Value > 0 Then
                    item.AddedLineCount = lineCount
                    addedTotal += lineCount.Value
                End If
            Next

            Return New GitNumstatTotals() With {
                .AddedLineCount = If(addedTotal > 0, CType(addedTotal, Integer?), Nothing),
                .RemovedLineCount = If(removedTotal > 0, CType(removedTotal, Integer?), Nothing)
            }
        End Function

        Private Shared Function CountFileLinesOrNothing(filePath As String) As Integer?
            Try
                If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
                    Return Nothing
                End If

                Dim lineCount = 0
                For Each lineText In File.ReadLines(filePath)
                    lineCount += 1
                Next

                If lineCount <= 0 Then
                    Return Nothing
                End If

                Return lineCount
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function ParseGitNumstatPerPath(outputText As String) As Dictionary(Of String, GitNumstatTotals)
            Dim result As New Dictionary(Of String, GitNumstatTotals)(StringComparer.OrdinalIgnoreCase)
            Dim normalized = If(outputText, String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim lines = normalized.Split({vbLf}, StringSplitOptions.None)

            For Each rawLine In lines
                Dim line = If(rawLine, String.Empty).TrimEnd()
                If String.IsNullOrWhiteSpace(line) Then
                    Continue For
                End If

                Dim parts = line.Split(ControlChars.Tab)
                If parts.Length < 3 Then
                    Continue For
                End If

                Dim filePath = parts(2).Trim()
                If String.IsNullOrWhiteSpace(filePath) Then
                    Continue For
                End If

                Dim addCount As Integer
                Dim removeCount As Integer
                Dim added As Integer? = Nothing
                Dim removed As Integer? = Nothing
                If Integer.TryParse(parts(0).Trim(), addCount) Then
                    added = Math.Max(0, addCount)
                End If
                If Integer.TryParse(parts(1).Trim(), removeCount) Then
                    removed = Math.Max(0, removeCount)
                End If

                result(filePath) = New GitNumstatTotals() With {
                    .AddedLineCount = If(added.HasValue AndAlso added.Value > 0, added, Nothing),
                    .RemovedLineCount = If(removed.HasValue AndAlso removed.Value > 0, removed, Nothing)
                }
            Next

            Return result
        End Function

        Private Shared Function CountGitNumstatTotals(outputText As String) As GitNumstatTotals
            Dim normalized = If(outputText, String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim lines = normalized.Split({vbLf}, StringSplitOptions.None)
            Dim added = 0
            Dim removed = 0

            For Each rawLine In lines
                Dim line = If(rawLine, String.Empty).TrimEnd()
                If String.IsNullOrWhiteSpace(line) Then
                    Continue For
                End If

                Dim parts = line.Split(ControlChars.Tab)
                If parts.Length < 3 Then
                    Continue For
                End If

                Dim addCount As Integer
                If Integer.TryParse(parts(0).Trim(), addCount) Then
                    added += Math.Max(0, addCount)
                End If

                Dim removeCount As Integer
                If Integer.TryParse(parts(1).Trim(), removeCount) Then
                    removed += Math.Max(0, removeCount)
                End If
            Next

            Return New GitNumstatTotals() With {
                .AddedLineCount = If(added > 0, CType(added, Integer?), Nothing),
                .RemovedLineCount = If(removed > 0, CType(removed, Integer?), Nothing)
            }
        End Function

        Private Shared Function ParseGitCommits(outputText As String) As List(Of GitCommitListEntry)
            Dim commits As New List(Of GitCommitListEntry)()
            Dim normalized = If(outputText, String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim lines = normalized.Split({vbLf}, StringSplitOptions.None)
            Dim separator = ChrW(&H1F)

            For Each rawLine In lines
                Dim line = If(rawLine, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(line) Then
                    Continue For
                End If

                Dim parts = line.Split({separator}, StringSplitOptions.None)
                Dim entry As New GitCommitListEntry() With {
                    .Sha = If(parts.Length > 0, parts(0).Trim(), String.Empty),
                    .ShortSha = If(parts.Length > 1, parts(1).Trim(), String.Empty),
                    .RelativeTime = If(parts.Length > 2, parts(2).Trim(), String.Empty),
                    .Subject = If(parts.Length > 3, parts(3).Trim(), String.Empty),
                    .Decorations = If(parts.Length > 4, parts(4).Trim(), String.Empty)
                }

                If String.IsNullOrWhiteSpace(entry.Sha) Then
                    Continue For
                End If

                commits.Add(entry)
            Next

            Return commits
        End Function

        Private Shared Function ParseGitBranches(outputText As String) As List(Of GitBranchListEntry)
            Dim branches As New List(Of GitBranchListEntry)()
            Dim normalized = If(outputText, String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim lines = normalized.Split({vbLf}, StringSplitOptions.None)
            Dim separator = ChrW(&H1F)

            For Each rawLine In lines
                Dim line = If(rawLine, String.Empty)
                If String.IsNullOrWhiteSpace(line) Then
                    Continue For
                End If

                Dim parts = line.Split({separator}, StringSplitOptions.None)
                If parts.Length < 2 Then
                    Continue For
                End If

                Dim headMarker = parts(0).Trim()
                Dim branchName = parts(1).Trim()
                If String.IsNullOrWhiteSpace(branchName) Then
                    Continue For
                End If

                branches.Add(New GitBranchListEntry() With {
                    .IsCurrent = StringComparer.Ordinal.Equals(headMarker, "*"),
                    .Name = branchName,
                    .IsRemote = branchName.StartsWith("origin/", StringComparison.OrdinalIgnoreCase) OrElse branchName.StartsWith("remotes/", StringComparison.OrdinalIgnoreCase),
                    .CommitShortSha = If(parts.Length > 2, parts(2).Trim(), String.Empty),
                    .RelativeTime = If(parts.Length > 3, parts(3).Trim(), String.Empty),
                    .Subject = If(parts.Length > 4, parts(4).Trim(), String.Empty)
                })
            Next

            branches.Sort(
                Function(a, b)
                    If a Is Nothing AndAlso b Is Nothing Then Return 0
                    If a Is Nothing Then Return 1
                    If b Is Nothing Then Return -1
                    If a.IsCurrent <> b.IsCurrent Then
                        Return If(a.IsCurrent, -1, 1)
                    End If

                    Return StringComparer.OrdinalIgnoreCase.Compare(a.Name, b.Name)
                End Function)

            Return branches
        End Function

        Private Shared Function BuildGitFileDiffPreview(repoRoot As String, entry As GitChangedFileListEntry) As String
            If String.IsNullOrWhiteSpace(repoRoot) OrElse entry Is Nothing OrElse String.IsNullOrWhiteSpace(entry.FilePath) Then
                Return "No file selected."
            End If

            If entry.IsUntracked Then
                Return BuildUntrackedFilePreview(repoRoot, entry.FilePath)
            End If

            Dim quotedPath = QuoteProcessArgument(entry.FilePath)
            Dim diffResult = RunProcessCapture("git", $"--no-pager diff --no-color HEAD -- {quotedPath}", repoRoot)
            If diffResult.ExitCode = 0 Then
                Dim output = NormalizePanelMultiline(diffResult.OutputText, String.Empty)
                If Not String.IsNullOrWhiteSpace(output) Then
                    Return output
                End If
            End If

            Dim fallbackResult = RunProcessCapture("git", $"--no-pager diff --no-color -- {quotedPath}", repoRoot)
            If fallbackResult.ExitCode = 0 Then
                Dim output = NormalizePanelMultiline(fallbackResult.OutputText, String.Empty)
                If Not String.IsNullOrWhiteSpace(output) Then
                    Return output
                End If
            End If

            Return "No diff output available for this file."
        End Function

        Private Shared Function BuildGitCommitPreview(repoRoot As String, commitSha As String) As String
            If String.IsNullOrWhiteSpace(repoRoot) OrElse String.IsNullOrWhiteSpace(commitSha) Then
                Return "No commit selected."
            End If

            Dim result = RunProcessCapture("git", $"--no-pager show --no-color --stat --patch --format=medium {QuoteProcessArgument(commitSha)}", repoRoot)
            If result.ExitCode <> 0 Then
                Return NormalizeProcessError(result)
            End If

            Return NormalizePanelMultiline(result.OutputText, "No commit preview available.")
        End Function

        Private Shared Function BuildGitBranchPreview(repoRoot As String, branchName As String) As String
            If String.IsNullOrWhiteSpace(repoRoot) OrElse String.IsNullOrWhiteSpace(branchName) Then
                Return "No branch selected."
            End If

            Dim historyResult = RunProcessCapture("git", $"--no-pager log --oneline --decorate -n 15 {QuoteProcessArgument(branchName)}", repoRoot)
            If historyResult.ExitCode <> 0 Then
                Return NormalizeProcessError(historyResult)
            End If

            Dim header = $"Recent history for {branchName}"
            Dim body = NormalizePanelMultiline(historyResult.OutputText, "No history available.")
            Return $"{header}{Environment.NewLine}{Environment.NewLine}{body}"
        End Function

        Private Shared Function BuildUntrackedFilePreview(repoRoot As String, relativePath As String) As String
            Try
                Dim fullPath = Path.Combine(repoRoot, relativePath)
                If Not File.Exists(fullPath) Then
                    Return "Untracked file preview unavailable (file not found)."
                End If

                Dim fileInfo As New FileInfo(fullPath)
                If fileInfo.Length > 512 * 1024 Then
                    Return $"Untracked file: {relativePath}{Environment.NewLine}{Environment.NewLine}(Preview skipped: file is larger than 512 KB)"
                End If

                Dim text = File.ReadAllText(fullPath)
                Dim normalized = NormalizePanelMultiline(text, String.Empty)
                If normalized.Length > 12000 Then
                    normalized = normalized.Substring(0, 12000) & Environment.NewLine & "... (truncated)"
                End If

                Return $"Untracked file: {relativePath}{Environment.NewLine}{Environment.NewLine}{normalized}"
            Catch ex As Exception
                Return $"Untracked file preview unavailable: {ex.Message}"
            End Try
        End Function

        Private Shared Function QuoteProcessArgument(value As String) As String
            Dim text = If(value, String.Empty)
            Return """" & text & """"
        End Function

        Private Shared Function FirstNonEmptyLine(value As String) As String
            Dim normalized = If(value, String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim lines = normalized.Split({vbLf}, StringSplitOptions.None)
            For Each line In lines
                Dim trimmed = If(line, String.Empty).Trim()
                If Not String.IsNullOrWhiteSpace(trimmed) Then
                    Return trimmed
                End If
            Next

            Return String.Empty
        End Function

        Private Shared Function NormalizePanelMultiline(value As String, fallback As String) As String
            Dim normalized = If(value, String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Trim()
            If String.IsNullOrWhiteSpace(normalized) Then
                Return If(fallback, String.Empty)
            End If

            Return normalized
        End Function

        Private Shared Function NormalizeProcessError(result As ProcessCaptureResult) As String
            Dim [error] = NormalizePanelMultiline(result.ErrorText, String.Empty)
            If Not String.IsNullOrWhiteSpace([error]) Then
                Return [error]
            End If

            Return NormalizePanelMultiline(result.OutputText, "No data available.")
        End Function

        Private Function ResolveQuickOpenWorkspaceCwd(destinationLabel As String) As String
            Dim source As String = String.Empty
            Dim targetCwd = NormalizeProjectPath(ResolveNewThreadTargetCwd(source))
            If String.IsNullOrWhiteSpace(targetCwd) Then
                ShowStatus($"No folder available to open in {destinationLabel}.", isError:=True, displayToast:=True)
                Return String.Empty
            End If

            If Not Directory.Exists(targetCwd) Then
                ShowStatus($"Folder not found: {targetCwd}", isError:=True, displayToast:=True)
                Return String.Empty
            End If

            Return targetCwd
        End Function

        Private Shared Function StartProcessInDirectory(fileName As String,
                                                        arguments As String,
                                                        workingDirectory As String,
                                                        Optional createNoWindow As Boolean = False) As Boolean
            Try
                Dim startInfo As New ProcessStartInfo(fileName) With {
                    .Arguments = If(arguments, String.Empty),
                    .WorkingDirectory = workingDirectory,
                    .UseShellExecute = False,
                    .CreateNoWindow = createNoWindow
                }
                Process.Start(startInfo)
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Function ResolveNewThreadTargetCwd(Optional ByRef sourceLabel As String = Nothing) As String
            Dim explicitTarget = NormalizeProjectPath(_newThreadTargetOverrideCwd)
            If Not String.IsNullOrWhiteSpace(explicitTarget) Then
                sourceLabel = "selected folder"
                Return explicitTarget
            End If

            Dim currentThreadTarget = NormalizeProjectPath(_currentThreadCwd)
            If Not String.IsNullOrWhiteSpace(currentThreadTarget) Then
                sourceLabel = "active thread"
                Return currentThreadTarget
            End If

            Dim workingDir = EffectiveThreadWorkingDirectory()
            If String.IsNullOrWhiteSpace(workingDir) Then
                workingDir = Environment.CurrentDirectory
            End If

            sourceLabel = "working dir"
            Return workingDir.Trim()
        End Function

        Private Sub SyncNewThreadTargetChip()
            Dim source As String = String.Empty
            Dim targetCwd = ResolveNewThreadTargetCwd(source)
            If String.IsNullOrWhiteSpace(targetCwd) Then
                _viewModel.SidebarNewThreadButtonText = "New thread"
                _viewModel.SidebarNewThreadToolTip = Nothing
                Return
            End If

            Dim folderLabel = BuildProjectGroupLabel(targetCwd)
            _viewModel.SidebarNewThreadButtonText = $"New thread for {folderLabel}"
            _viewModel.SidebarNewThreadToolTip = $"{source}: {targetCwd}"
        End Sub

        Private Function ShowNewThreadFolderPicker() As String
            Dim initialDirectory = ResolveNewThreadTargetCwd()
            Dim dialog As New Microsoft.Win32.OpenFolderDialog() With {
                .Title = "Choose folder for new thread"
            }

            If Not String.IsNullOrWhiteSpace(initialDirectory) AndAlso Directory.Exists(initialDirectory) Then
                dialog.InitialDirectory = initialDirectory
            End If

            Dim result = dialog.ShowDialog(Me)
            If Not result.HasValue OrElse Not result.Value Then
                Return String.Empty
            End If

            Return If(dialog.FolderName, String.Empty).Trim()
        End Function

        Private Async Function ChooseFolderAndStartNewThreadAsync() As Task
            If Not _viewModel.IsSidebarNewThreadEnabled Then
                Return
            End If

            Dim chosenFolder = ShowNewThreadFolderPicker()
            If String.IsNullOrWhiteSpace(chosenFolder) Then
                Return
            End If

            _newThreadTargetOverrideCwd = chosenFolder
            SyncNewThreadTargetChip()
            Await StartThreadAsync()
        End Function

        Private Sub MainWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles Me.Closing
            _toastTimer.Stop()
            _watchdogTimer.Stop()
            _reconnectUiTimer.Stop()
            CancelReconnect()
            CancelActiveThreadSelectionLoad()
            SaveSettings()

            ShutdownClientForAppClose()
        End Sub

        Private Sub ShutdownClientForAppClose()
            BeginUserDisconnectSessionTransition()

            Dim client = DetachCurrentClient()
            If client Is Nothing Then
                Return
            End If

            _disconnecting = True
            Try
                ' Avoid deadlocking the UI thread during window close by not synchronously blocking
                ' on DisconnectAsync(), which captures the WPF synchronization context.
                RemoveHandler client.RawMessage, AddressOf ClientOnRawMessage
                RemoveHandler client.NotificationReceived, AddressOf ClientOnNotification
                RemoveHandler client.ServerRequestReceived, AddressOf ClientOnServerRequest
                RemoveHandler client.Disconnected, AddressOf ClientOnDisconnected

                client.StopAsync("Application closing.").ConfigureAwait(False).GetAwaiter().GetResult()
            Catch
            Finally
                _disconnecting = False
            End Try
        End Sub

        Private Sub MainWindow_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Me.PreviewKeyDown
            If Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.Enter Then
                If TryExecuteShellCommand(_viewModel.ShellSendCommand) Then
                    e.Handled = True
                End If
                Return
            End If

            If e.Key = Key.F5 Then
                If TryExecuteShellCommand(_viewModel.ShellRefreshThreadsCommand) Then
                    e.Handled = True
                End If
                Return
            End If

            If e.Key = Key.F6 Then
                If TryExecuteShellCommand(_viewModel.ShellRefreshModelsCommand) Then
                    e.Handled = True
                End If
                Return
            End If

            If Keyboard.Modifiers = (ModifierKeys.Control Or ModifierKeys.Shift) AndAlso e.Key = Key.N Then
                If TryExecuteShellCommand(_viewModel.ShellNewThreadCommand) Then
                    e.Handled = True
                End If
                Return
            End If

            If Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.F Then
                e.Handled = TryExecuteShellCommand(_viewModel.ShellFocusThreadSearchCommand)
                Return
            End If

            If Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.OemComma Then
                e.Handled = TryExecuteShellCommand(_viewModel.ShellOpenSettingsCommand)
            End If
        End Sub

        Private Shared Function TryExecuteShellCommand(command As ICommand) As Boolean
            If command Is Nothing Then
                Return False
            End If

            If Not command.CanExecute(Nothing) Then
                Return False
            End If

            command.Execute(Nothing)
            Return True
        End Function

        Private Sub RefreshCommandCanExecuteStates()
            CommandManager.InvalidateRequerySuggested()

            If _sessionCoordinator IsNot Nothing Then
                _sessionCoordinator.RefreshCommandCanExecuteStates()
            End If
        End Sub

        Private Shared Sub RaiseAsyncCommandCanExecuteChanged(command As ICommand)
            Dim asyncCommand = TryCast(command, AsyncRelayCommand)
            If asyncCommand Is Nothing Then
                Return
            End If

            asyncCommand.RaiseCanExecuteChanged()
        End Sub

        Private Async Function RunUiActionAsync(operation As Func(Of Task)) As Task
            Try
                Await operation().ConfigureAwait(True)
            Catch ex As TimeoutException
                ShowStatus($"Request timed out: {ex.Message}", isError:=True, displayToast:=True)
                AppendSystemMessage($"Timeout: {ex.Message}")
            Catch ex As Exception
                ShowStatus($"Error: {ex.Message}", isError:=True, displayToast:=True)
                AppendSystemMessage($"Error: {ex.Message}")
            End Try
        End Function

        Private Sub RunOnUi(action As Action)
            If action Is Nothing Then
                Return
            End If

            If Dispatcher.CheckAccess() Then
                action()
            Else
                Dispatcher.BeginInvoke(action)
            End If
        End Sub

        Private Function RunOnUiAsync(action As Func(Of Task)) As Task
            Dim tcs As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)

            RunOnUi(
                Async Sub()
                    Try
                        Await action().ConfigureAwait(True)
                        tcs.TrySetResult(True)
                    Catch ex As Exception
                        tcs.TrySetException(ex)
                    End Try
                End Sub)

            Return tcs.Task
        End Function

        Private Sub FireAndForget(task As Task)
            task.ContinueWith(
                Sub(faulted)
                    Dim [error] = faulted.Exception
                    If [error] Is Nothing Then
                        Return
                    End If

                    RunOnUi(Sub() AppendSystemMessage($"Background error: {[error].GetBaseException().Message}"))
                End Sub,
                TaskContinuationOptions.OnlyOnFaulted)
        End Sub

        Private Sub InitializeStatusUi()
            _toastTimer.Interval = TimeSpan.FromMilliseconds(3500)
            AddHandler _toastTimer.Tick, AddressOf OnToastTimerTick
            _workspaceHintOverlayTimer.Interval = TimeSpan.FromMilliseconds(8000)
            AddHandler _workspaceHintOverlayTimer.Tick, AddressOf OnWorkspaceHintOverlayTimerTick
            _viewModel.StatusText = "Ready."
            StatusBarPaneHost.LblStatus.Foreground = ResolveBrush("TextPrimaryBrush", Brushes.Black)
        End Sub

        Private Sub ShowStatus(message As String,
                               Optional isError As Boolean = False,
                               Optional displayToast As Boolean = False)
            _viewModel.StatusText = message
            StatusBarPaneHost.LblStatus.Foreground = If(isError,
                                      ResolveBrush("DangerBrush", Brushes.DarkRed),
                                      ResolveBrush("TextPrimaryBrush", Brushes.Black))

            Dim suppressHintToast = displayToast AndAlso
                                    Not isError AndAlso
                                    _viewModel IsNot Nothing AndAlso
                                    _viewModel.SettingsPanel IsNot Nothing AndAlso
                                    _viewModel.SettingsPanel.DisableConnectionInitializedToast

            If displayToast AndAlso Not suppressHintToast Then
                ShowToast(message, isError)
            End If
        End Sub

        Private Sub ShowToast(message As String, isError As Boolean)
            If String.IsNullOrWhiteSpace(message) Then
                Return
            End If

            LblToastText.Text = message
            If isError Then
                ToastOverlay.Background = ResolveBrush("ToastErrorBackgroundBrush", New SolidColorBrush(Color.FromRgb(255, 244, 242)))
                ToastOverlay.BorderBrush = ResolveBrush("ToastErrorBorderBrush", New SolidColorBrush(Color.FromRgb(242, 184, 177)))
                LblToastText.Foreground = ResolveBrush("DangerBrush", Brushes.DarkRed)
            Else
                ToastOverlay.Background = ResolveBrush("ToastInfoBackgroundBrush", New SolidColorBrush(Color.FromRgb(238, 247, 255)))
                ToastOverlay.BorderBrush = ResolveBrush("ToastInfoBorderBrush", New SolidColorBrush(Color.FromRgb(181, 208, 244)))
                LblToastText.Foreground = ResolveBrush("TextPrimaryBrush", Brushes.Black)
            End If

            ToastOverlay.Visibility = Visibility.Visible
            _toastTimer.Stop()
            _toastTimer.Start()
        End Sub

        Private Sub HideToast()
            _toastTimer.Stop()
            If ToastOverlay IsNot Nothing Then
                ToastOverlay.Visibility = Visibility.Collapsed
            End If
        End Sub

        Private Function ResolveBrush(resourceKey As String, fallback As Brush) As Brush
            If String.IsNullOrWhiteSpace(resourceKey) Then
                Return fallback
            End If

            Dim resolved = TryCast(TryFindResource(resourceKey), Brush)
            Return If(resolved, fallback)
        End Function

        Private Sub OnToastTimerTick(sender As Object, e As EventArgs)
            HideToast()
        End Sub

        Private Sub OnWorkspaceHintOverlayTimerTick(sender As Object, e As EventArgs)
            _workspaceHintOverlayTimer.Stop()
            DismissWorkspaceHintOverlay()
        End Sub

        Private Sub RefreshControlStates()
            SyncSessionStateViewModel()
            Dim session = _viewModel.SessionState
            Dim connected = session.IsConnected
            Dim authenticated = session.IsConnectedAndAuthenticated

            _viewModel.SettingsPanel.CanExportDiagnostics = True
            RefreshTurnComposerAndShellControls(session, authenticated)
            RefreshApprovalControlState(authenticated)
            RefreshThreadPanelControlState(session)
            RefreshWorkspaceHintState(connected, authenticated)
            RefreshPostStateSyncUi()
            UpdateWorkspaceEmptyStateVisibility()
        End Sub

        Private Sub RefreshTurnComposerAndShellControls(session As SessionStateViewModel,
                                                        authenticated As Boolean)
            If session Is Nothing Then
                Return
            End If

            Dim hasActiveTurn = HasActiveRuntimeTurnForCurrentThread()
            If Not hasActiveTurn AndAlso session.HasCurrentTurn AndAlso Not RuntimeHasTurnHistoryForCurrentThread() Then
                hasActiveTurn = True
            End If
            Dim canUseExistingThreadTurnControls = authenticated AndAlso Not _threadContentLoading AndAlso session.HasCurrentThread
            Dim canStartTurn = authenticated AndAlso Not _threadContentLoading

            _viewModel.TurnComposer.CanStartTurn = canStartTurn AndAlso Not hasActiveTurn
            _viewModel.TurnComposer.CanSteerTurn = canUseExistingThreadTurnControls AndAlso hasActiveTurn
            _viewModel.TurnComposer.CanInterruptTurn = canUseExistingThreadTurnControls AndAlso hasActiveTurn
            _viewModel.TurnComposer.StartTurnVisibility = If(_viewModel.TurnComposer.CanInterruptTurn, Visibility.Collapsed, Visibility.Visible)
            _viewModel.TurnComposer.InterruptTurnVisibility = If(_viewModel.TurnComposer.CanInterruptTurn, Visibility.Visible, Visibility.Collapsed)
            _viewModel.TurnComposer.IsInputEnabled = authenticated
            _viewModel.TurnComposer.IsModelEnabled = authenticated
            _viewModel.TurnComposer.IsReasoningEnabled = authenticated
            _viewModel.TurnComposer.IsApprovalPolicyEnabled = authenticated
            _viewModel.TurnComposer.IsSandboxEnabled = authenticated

            _viewModel.IsSidebarNewThreadEnabled = authenticated AndAlso Not _threadContentLoading
            _viewModel.IsSidebarAutomationsEnabled = True
            _viewModel.IsSidebarSkillsEnabled = True
            _viewModel.IsSidebarSettingsEnabled = True
            _viewModel.IsSettingsBackEnabled = True
            _viewModel.IsQuickOpenVscEnabled = True
            _viewModel.IsQuickOpenTerminalEnabled = True
            _viewModel.TranscriptPanel.CollapseCommandDetailsByDefault = hasActiveTurn
        End Sub

        Private Sub RefreshApprovalControlState(authenticated As Boolean)
            Dim hasActiveApproval = _turnWorkflowCoordinator IsNot Nothing AndAlso _turnWorkflowCoordinator.HasActiveApproval
            _viewModel.ApprovalPanel.UpdateAvailability(authenticated, hasActiveApproval)
        End Sub

        Private Sub RefreshThreadPanelControlState(session As SessionStateViewModel)
            _viewModel.ThreadsPanel.UpdateInteractionState(session, _threadsLoading, _threadContentLoading)
            _viewModel.AreThreadLegacyFilterControlsEnabled = _viewModel.ThreadsPanel.IsSearchEnabled
            UpdateThreadsPanelStateHintBubbleVisibility()
        End Sub

        Private Sub RefreshWorkspaceHintState(connected As Boolean,
                                              authenticated As Boolean)
            _viewModel.WorkspaceHintText = BuildWorkspaceHint(connected, authenticated)
            UpdateWorkspaceHintOverlayVisibility()
        End Sub

        Private Sub RefreshPostStateSyncUi()
            UpdateRuntimeFieldState()
            RefreshCommandCanExecuteStates()
            SyncThreadToolbarMenus()
            SyncSidebarNewThreadMenu()
            SyncNewThreadTargetChip()
        End Sub

        Private Function BuildWorkspaceHint(connected As Boolean, authenticated As Boolean) As String
            If Not connected Then
                Return "Connect to Codex App Server from Settings to begin."
            End If

            If Not authenticated Then
                Return "Authentication required: sign in from Settings to unlock threads and turns."
            End If

            If String.IsNullOrWhiteSpace(_currentThreadId) Then
                Return "Select a thread from the left panel, or send your first instruction to start a new one."
            End If

            If String.IsNullOrWhiteSpace(_currentTurnId) Then
                Return "Ready. Send with Ctrl+Enter."
            End If

            Return "Turn in progress. Use Steer to refine or Interrupt to stop execution."
        End Function

        Private Sub ShowControlCenterTab()
            ShowSettingsView()
        End Sub

        Private Sub ShowWorkspaceSidebarTab()
            ShowWorkspaceView()
        End Sub

        Private Sub ShowThreadsSidebarTab()
            ShowWorkspaceView()
        End Sub

        Private Sub ShowSettingsView()
            _viewModel.SidebarMainViewVisibility = Visibility.Collapsed
            _viewModel.SidebarSettingsViewVisibility = Visibility.Visible
            UpdateSidebarSelectionState(showSettings:=True)
        End Sub

        Private Sub ShowWorkspaceView()
            _viewModel.SidebarSettingsViewVisibility = Visibility.Collapsed
            _viewModel.SidebarMainViewVisibility = Visibility.Visible
            UpdateSidebarSelectionState(showSettings:=False)
        End Sub

        Private Sub UpdateSidebarSelectionState(showSettings As Boolean)
            Dim newThreadTag As String
            If showSettings Then
                newThreadTag = String.Empty
            Else
                newThreadTag = "Active"
            End If

            _viewModel.SidebarNewThreadNavTag = newThreadTag
            _viewModel.SidebarSettingsNavTag = If(showSettings, "Active", String.Empty)
            _viewModel.SidebarAutomationsNavTag = String.Empty
            _viewModel.SidebarSkillsNavTag = String.Empty
        End Sub
        Private Async Function ExportDiagnosticsAsync() As Task
            SaveSettings()
            SyncSessionStateViewModel()

            Dim diagnosticsRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                                               "CodexNativeAgentDiagnostics")
            Directory.CreateDirectory(diagnosticsRoot)

            Dim stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss")
            Dim outputFolder = Path.Combine(diagnosticsRoot, $"diag-{stamp}")
            Directory.CreateDirectory(outputFolder)

            File.WriteAllText(Path.Combine(outputFolder, "transcript.log"), _viewModel.TranscriptPanel.FullTranscriptText)
            File.WriteAllText(Path.Combine(outputFolder, "protocol.log"), _viewModel.TranscriptPanel.FullProtocolText)
            File.WriteAllText(Path.Combine(outputFolder, "approval.txt"), _viewModel.ApprovalPanel.SummaryText)
            File.WriteAllText(Path.Combine(outputFolder, "rate-limits.txt"), _viewModel.SettingsPanel.RateLimitsText)

            Dim threadBuilder As New StringBuilder()
            For Each entry In _threadEntries
                threadBuilder.AppendLine($"{entry.Id}{ControlChars.Tab}{entry.LastActiveAt}{ControlChars.Tab}{entry.Cwd}{ControlChars.Tab}{entry.Preview}")
            Next
            File.WriteAllText(Path.Combine(outputFolder, "threads.tsv"), threadBuilder.ToString())

            Dim settingsObject As New JsonObject()
            settingsObject("codexPath") = _viewModel.SettingsPanel.CodexPath.Trim()
            settingsObject("serverArgs") = _viewModel.SettingsPanel.ServerArgs.Trim()
            settingsObject("workingDir") = _viewModel.SettingsPanel.WorkingDir.Trim()
            settingsObject("windowsCodexHome") = _viewModel.SettingsPanel.WindowsCodexHome.Trim()
            settingsObject("rememberApiKey") = _viewModel.SettingsPanel.RememberApiKey
            settingsObject("autoLoginApiKey") = _viewModel.SettingsPanel.AutoLoginApiKey
            settingsObject("autoReconnect") = _viewModel.SettingsPanel.AutoReconnect
            settingsObject("disableWorkspaceHintOverlay") = _viewModel.SettingsPanel.DisableWorkspaceHintOverlay
            settingsObject("disableConnectionInitializedToast") = _viewModel.SettingsPanel.DisableConnectionInitializedToast
            settingsObject("disableThreadsPanelHints") = _viewModel.SettingsPanel.DisableThreadsPanelHints
            settingsObject("showEventDotsInTranscript") = _viewModel.SettingsPanel.ShowEventDotsInTranscript
            settingsObject("showSystemDotsInTranscript") = _viewModel.SettingsPanel.ShowSystemDotsInTranscript
            settingsObject("theme") = _currentTheme
            settingsObject("density") = _currentDensity
            settingsObject("apiKeyMasked") = MaskSecret(_viewModel.SettingsPanel.ApiKey.Trim(), 4)
            settingsObject("externalIdTokenMasked") = MaskSecret(_viewModel.SettingsPanel.ExternalIdToken.Trim(), 6)
            settingsObject("externalAccessTokenMasked") = MaskSecret(_viewModel.SettingsPanel.ExternalAccessToken.Trim(), 6)

            Dim connectionObject As New JsonObject()
            connectionObject("isConnected") = _viewModel.SessionState.IsConnected
            connectionObject("isAuthenticated") = _viewModel.SessionState.IsAuthenticated
            connectionObject("expectedConnection") = _viewModel.SessionState.ConnectionExpected
            connectionObject("reconnectInProgress") = _viewModel.SessionState.IsReconnectInProgress
            connectionObject("reconnectAttempt") = _viewModel.SessionState.ReconnectAttempt
            connectionObject("currentLoginId") = _viewModel.SessionState.CurrentLoginId
            connectionObject("currentThreadId") = _viewModel.SessionState.CurrentThreadId
            connectionObject("currentTurnId") = _viewModel.SessionState.CurrentTurnId
            connectionObject("lastActivityUtc") = _viewModel.SessionState.LastActivityUtc.ToString("O")
            connectionObject("processId") = _viewModel.SessionState.ProcessId
            connectionObject("nextReconnectAttemptUtc") = If(_viewModel.SessionState.NextReconnectAttemptUtc.HasValue,
                                                             _viewModel.SessionState.NextReconnectAttemptUtc.Value.ToString("O"),
                                                             String.Empty)

            Dim uiObject As New JsonObject()
            uiObject("threadSearch") = _viewModel.ThreadsPanel.SearchText
            uiObject("threadSort") = ThreadSortLabel(_viewModel.ThreadsPanel.SortIndex)
            uiObject("model") = _viewModel.TurnComposer.SelectedModelId
            uiObject("approvalPolicy") = _viewModel.TurnComposer.SelectedApprovalPolicy
            uiObject("sandbox") = _viewModel.TurnComposer.SelectedSandbox
            uiObject("reasoningEffort") = _viewModel.TurnComposer.SelectedReasoningEffort
            uiObject("approvalSummary") = _viewModel.ApprovalPanel.SummaryText
            uiObject("approvalPendingQueueCount") = _viewModel.ApprovalPanel.PendingQueueCount
            uiObject("approvalActiveMethodName") = _viewModel.ApprovalPanel.ActiveMethodName
            uiObject("approvalLastQueuedMethodName") = _viewModel.ApprovalPanel.LastQueuedMethodName
            uiObject("approvalLastResolvedAction") = _viewModel.ApprovalPanel.LastResolvedAction
            uiObject("approvalLastResolvedDecision") = _viewModel.ApprovalPanel.LastResolvedDecision
            uiObject("approvalLastError") = _viewModel.ApprovalPanel.LastErrorText
            uiObject("approvalLastQueueUpdatedUtc") = If(_viewModel.ApprovalPanel.LastQueueUpdatedUtc.HasValue,
                                                         _viewModel.ApprovalPanel.LastQueueUpdatedUtc.Value.ToString("O"),
                                                         String.Empty)
            uiObject("approvalLastResolvedUtc") = If(_viewModel.ApprovalPanel.LastResolvedUtc.HasValue,
                                                     _viewModel.ApprovalPanel.LastResolvedUtc.Value.ToString("O"),
                                                     String.Empty)
            uiObject("threadsLoading") = _viewModel.ThreadsPanel.IsLoading
            uiObject("threadContentLoading") = _viewModel.ThreadsPanel.IsThreadContentLoading
            uiObject("threadRefreshError") = _viewModel.ThreadsPanel.RefreshErrorText
            uiObject("threadRefreshLastCount") = _viewModel.ThreadsPanel.LastRefreshThreadCount
            uiObject("threadRefreshLastStartedUtc") = If(_viewModel.ThreadsPanel.LastRefreshStartedUtc.HasValue,
                                                         _viewModel.ThreadsPanel.LastRefreshStartedUtc.Value.ToString("O"),
                                                         String.Empty)
            uiObject("threadRefreshLastCompletedUtc") = If(_viewModel.ThreadsPanel.LastRefreshCompletedUtc.HasValue,
                                                           _viewModel.ThreadsPanel.LastRefreshCompletedUtc.Value.ToString("O"),
                                                           String.Empty)
            uiObject("threadSelectionLoadVersion") = _viewModel.ThreadsPanel.SelectionLoadVersion
            uiObject("threadSelectionLoadThreadId") = _viewModel.ThreadsPanel.SelectionLoadThreadId
            uiObject("threadSelectionHasActiveLoad") = _viewModel.ThreadsPanel.HasActiveSelectionLoad
            uiObject("threadSelectionLastError") = _viewModel.ThreadsPanel.LastSelectionLoadErrorText
            uiObject("threadSelectionLastStartedUtc") = If(_viewModel.ThreadsPanel.LastSelectionLoadStartedUtc.HasValue,
                                                           _viewModel.ThreadsPanel.LastSelectionLoadStartedUtc.Value.ToString("O"),
                                                           String.Empty)
            uiObject("threadSelectionLastCompletedUtc") = If(_viewModel.ThreadsPanel.LastSelectionLoadCompletedUtc.HasValue,
                                                             _viewModel.ThreadsPanel.LastSelectionLoadCompletedUtc.Value.ToString("O"),
                                                             String.Empty)
            uiObject("theme") = _currentTheme
            uiObject("density") = _currentDensity

            Dim snapshot As New JsonObject()
            snapshot("generatedAtLocal") = DateTime.Now.ToString("O")
            snapshot("generatedAtUtc") = DateTimeOffset.UtcNow.ToString("O")
            snapshot("statusText") = StatusBarPaneHost.LblStatus.Text
            snapshot("settings") = settingsObject
            snapshot("connection") = connectionObject
            snapshot("ui") = uiObject

            File.WriteAllText(Path.Combine(outputFolder, "config-snapshot.json"), snapshot.ToJsonString(_settingsJsonOptions))

            Dim zipPath = outputFolder & ".zip"
            If File.Exists(zipPath) Then
                File.Delete(zipPath)
            End If

            ZipFile.CreateFromDirectory(outputFolder, zipPath, CompressionLevel.Optimal, includeBaseDirectory:=False)

            AppendSystemMessage($"Diagnostics exported: {zipPath}")
            ShowStatus($"Diagnostics exported: {zipPath}", displayToast:=True)
            Await Task.CompletedTask
        End Function

        Private Shared Function MaskSecret(value As String, visibleChars As Integer) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return String.Empty
            End If

            Dim trimmed = value.Trim()
            If trimmed.Length <= visibleChars * 2 Then
                Return New String("*"c, trimmed.Length)
            End If

            Dim prefix = trimmed.Substring(0, visibleChars)
            Dim suffix = trimmed.Substring(trimmed.Length - visibleChars)
            Return $"{prefix}{New String("*"c, trimmed.Length - (visibleChars * 2))}{suffix}"
        End Function

        Private Shared Sub ScrollTextBoxToBottom(textBox As TextBox)
            If textBox Is Nothing Then
                Return
            End If

            textBox.CaretIndex = textBox.Text.Length
            textBox.ScrollToEnd()
        End Sub

        Private Sub OnTranscriptScrollChanged(sender As Object, e As ScrollChangedEventArgs)
            If e Is Nothing OrElse _suppressTranscriptScrollTracking Then
                Return
            End If

            Dim scroller = TryCast(e.OriginalSource, ScrollViewer)
            If scroller Is Nothing Then
                Dim root = TryCast(sender, DependencyObject)
                If root Is Nothing Then
                    Return
                End If

                scroller = FindVisualDescendant(Of ScrollViewer)(root)
            End If

            If scroller Is Nothing Then
                Return
            End If

            If Not IsTurnInProgressForTranscriptAutoScroll() Then
                _transcriptAutoScrollEnabled = True
                Return
            End If

            If IsScrollViewerNearBottom(scroller) Then
                _transcriptAutoScrollEnabled = True
                Return
            End If

            If e.VerticalChange < 0 Then
                _transcriptAutoScrollEnabled = False
            End If
        End Sub

        Private Function IsTurnInProgressForTranscriptAutoScroll() As Boolean
            Dim session = If(_viewModel, Nothing)?.SessionState
            Return session IsNot Nothing AndAlso session.HasCurrentTurn
        End Function

        Private Function IsWorkspaceHintOverlayContext() As Boolean
            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.WorkspaceHintOverlay Is Nothing Then
                Return False
            End If

            Dim hintText = If(_viewModel, Nothing)?.WorkspaceHintText
            If String.IsNullOrWhiteSpace(hintText) Then
                Return False
            End If

            Dim transcriptPanel = If(_viewModel, Nothing)?.TranscriptPanel
            If transcriptPanel IsNot Nothing AndAlso transcriptPanel.LoadingOverlayVisibility = Visibility.Visible Then
                Return False
            End If

            Dim session = If(_viewModel, Nothing)?.SessionState
            If session Is Nothing Then
                Return False
            End If

            Return session.IsConnectedAndAuthenticated AndAlso
                   Not session.HasCurrentThread AndAlso
                   Not session.HasCurrentTurn
        End Function

        Private Shared Function IsThreadsPanelHintStateText(stateText As String) As Boolean
            Dim normalized = If(stateText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalized) Then
                Return False
            End If

            Select Case normalized
                Case "No threads loaded yet.",
                     "Connect to Codex App Server to load threads.",
                     "Authentication required. Sign in to start or load threads.",
                     "No threads found. Start a new thread to begin.",
                     "All project folders are collapsed. Expand a folder to view threads.",
                     "No threads match the current search/filter."
                    Return True
            End Select

            Return False
        End Function

        Private Sub UpdateThreadsPanelStateHintBubbleVisibility()
            If SidebarPaneHost Is Nothing Then
                Return
            End If

            Dim bubble = SidebarPaneHost.ThreadsStateHintBubble
            Dim dismissButton = SidebarPaneHost.BtnDismissThreadsStateHint
            If bubble Is Nothing OrElse dismissButton Is Nothing Then
                Return
            End If

            Dim stateText = If(_viewModel, Nothing)?.ThreadsPanel?.StateText
            Dim hasStateText = Not String.IsNullOrWhiteSpace(stateText)
            If Not hasStateText Then
                bubble.Visibility = Visibility.Collapsed
                dismissButton.Visibility = Visibility.Collapsed
                _threadsPanelHintBubbleHintKey = String.Empty
                _threadsPanelHintBubbleDismissedForCurrentHint = False
                Return
            End If

            Dim bubbleDisabled = _viewModel IsNot Nothing AndAlso
                                 _viewModel.SettingsPanel IsNot Nothing AndAlso
                                 _viewModel.SettingsPanel.DisableThreadsPanelHints
            If bubbleDisabled Then
                bubble.Visibility = Visibility.Collapsed
                dismissButton.Visibility = Visibility.Collapsed
                Return
            End If

            Dim trimmedStateText = stateText.Trim()
            Dim isHint = IsThreadsPanelHintStateText(trimmedStateText)
            dismissButton.Visibility = If(isHint, Visibility.Visible, Visibility.Collapsed)

            If Not isHint Then
                bubble.Visibility = Visibility.Visible
                _threadsPanelHintBubbleHintKey = String.Empty
                _threadsPanelHintBubbleDismissedForCurrentHint = False
                Return
            End If

            Dim hintChanged = Not StringComparer.Ordinal.Equals(trimmedStateText, _threadsPanelHintBubbleHintKey)
            If hintChanged Then
                _threadsPanelHintBubbleHintKey = trimmedStateText
                _threadsPanelHintBubbleDismissedForCurrentHint = False
            End If

            If _threadsPanelHintBubbleDismissedForCurrentHint Then
                bubble.Visibility = Visibility.Collapsed
                Return
            End If

            bubble.Visibility = Visibility.Visible
        End Sub

        Private Sub UpdateWorkspaceHintOverlayVisibility()
            If WorkspacePaneHost Is Nothing Then
                Return
            End If

            Dim banner = WorkspacePaneHost.WorkspaceHintBanner
            Dim overlay = WorkspacePaneHost.WorkspaceHintOverlay
            If banner Is Nothing OrElse overlay Is Nothing Then
                Return
            End If

            Dim hintText = If(_viewModel, Nothing)?.WorkspaceHintText
            Dim hasHintText = Not String.IsNullOrWhiteSpace(hintText)
            Dim overlayContext = IsWorkspaceHintOverlayContext()
            Dim hintsDisabled = _viewModel IsNot Nothing AndAlso
                                _viewModel.SettingsPanel IsNot Nothing AndAlso
                                _viewModel.SettingsPanel.DisableWorkspaceHintOverlay
            Dim overlayAllowed = overlayContext AndAlso Not hintsDisabled

            banner.Visibility = If(hasHintText AndAlso Not overlayContext AndAlso Not hintsDisabled,
                                   Visibility.Visible,
                                   Visibility.Collapsed)

            If Not overlayAllowed Then
                _workspaceHintOverlayTimer.Stop()
                overlay.Visibility = Visibility.Collapsed
                If Not overlayContext Then
                    _workspaceHintOverlayHintKey = String.Empty
                    _workspaceHintOverlayDismissedForCurrentHint = False
                End If
                Return
            End If

            Dim hintKey = hintText.Trim()
            Dim hintChanged = Not StringComparer.Ordinal.Equals(hintKey, _workspaceHintOverlayHintKey)
            If hintChanged Then
                _workspaceHintOverlayHintKey = hintKey
                _workspaceHintOverlayDismissedForCurrentHint = False
            End If

            If _workspaceHintOverlayDismissedForCurrentHint Then
                overlay.Visibility = Visibility.Collapsed
                _workspaceHintOverlayTimer.Stop()
                Return
            End If

            If overlay.Visibility <> Visibility.Visible OrElse hintChanged Then
                overlay.Visibility = Visibility.Visible
                _workspaceHintOverlayTimer.Stop()
                _workspaceHintOverlayTimer.Start()
            End If
        End Sub

        Private Sub DismissWorkspaceHintOverlay(Optional dismissCurrentHint As Boolean = True)
            _workspaceHintOverlayTimer.Stop()
            If dismissCurrentHint Then
                _workspaceHintOverlayDismissedForCurrentHint = True
            End If

            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.WorkspaceHintOverlay Is Nothing Then
                Return
            End If

            WorkspacePaneHost.WorkspaceHintOverlay.Visibility = Visibility.Collapsed
        End Sub

        Private Sub DismissThreadsPanelStateHint(Optional dismissCurrentHint As Boolean = True)
            If dismissCurrentHint Then
                _threadsPanelHintBubbleDismissedForCurrentHint = True
            End If

            If SidebarPaneHost Is Nothing OrElse SidebarPaneHost.ThreadsStateHintBubble Is Nothing Then
                Return
            End If

            SidebarPaneHost.ThreadsStateHintBubble.Visibility = Visibility.Collapsed
        End Sub

        Private Sub UpdateWorkspaceEmptyStateVisibility()
            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.ImgTimeToBuildEmptyState Is Nothing Then
                Return
            End If

            Dim transcriptPanel = If(_viewModel, Nothing)?.TranscriptPanel
            Dim hasVisibleTranscriptEntries = False
            If transcriptPanel IsNot Nothing AndAlso transcriptPanel.Items IsNot Nothing Then
                For Each entry In transcriptPanel.Items
                    Dim transcriptEntry = TryCast(entry, ViewModels.Transcript.TranscriptEntryViewModel)
                    If transcriptEntry Is Nothing Then
                        Continue For
                    End If

                    If transcriptEntry.RowVisibility = Visibility.Visible Then
                        hasVisibleTranscriptEntries = True
                        Exit For
                    End If
                Next
            End If
            Dim transcriptLoading = transcriptPanel IsNot Nothing AndAlso
                                    transcriptPanel.LoadingOverlayVisibility = Visibility.Visible
            Dim noThreadSelected = String.IsNullOrWhiteSpace(_currentThreadId)
            Dim noActiveTurn = String.IsNullOrWhiteSpace(_currentTurnId)
            Dim showDraftNewThreadEmptyState = _pendingNewThreadFirstPromptSelection
            Dim showEmptyState = (noThreadSelected OrElse showDraftNewThreadEmptyState) AndAlso
                                 noActiveTurn AndAlso
                                 Not hasVisibleTranscriptEntries AndAlso
                                 Not transcriptLoading

            WorkspacePaneHost.ImgTimeToBuildEmptyState.Visibility = If(showEmptyState, Visibility.Visible, Visibility.Collapsed)
        End Sub

        Private Shared Function IsScrollViewerNearBottom(scroller As ScrollViewer) As Boolean
            If scroller Is Nothing Then
                Return True
            End If

            Const bottomEpsilon As Double = 2.0
            Return scroller.VerticalOffset >= Math.Max(0, scroller.ScrollableHeight - bottomEpsilon)
        End Function

        Private Sub ScrollTranscriptToBottom()
            If WorkspacePaneHost Is Nothing Then
                Return
            End If

            UpdateWorkspaceEmptyStateVisibility()

            If Not IsTurnInProgressForTranscriptAutoScroll() Then
                _transcriptAutoScrollEnabled = True
            ElseIf Not _transcriptAutoScrollEnabled Then
                Return
            End If

            Dim transcriptList = WorkspacePaneHost.LstTranscript
            If transcriptList IsNot Nothing Then
                Dim scroller = FindVisualDescendant(Of ScrollViewer)(transcriptList)
                If scroller IsNot Nothing Then
                    _suppressTranscriptScrollTracking = True
                    Try
                        scroller.ScrollToBottom()
                    Finally
                        _suppressTranscriptScrollTracking = False
                    End Try
                ElseIf transcriptList.Items IsNot Nothing AndAlso transcriptList.Items.Count > 0 Then
                    transcriptList.ScrollIntoView(transcriptList.Items(transcriptList.Items.Count - 1))
                End If
            End If

            ScrollTextBoxToBottom(WorkspacePaneHost.TxtTranscript)
        End Sub

        Private Shared Function FindVisualDescendant(Of T As DependencyObject)(root As DependencyObject) As T
            If root Is Nothing Then
                Return Nothing
            End If

            Dim childCount = VisualTreeHelper.GetChildrenCount(root)
            For i = 0 To childCount - 1
                Dim child = VisualTreeHelper.GetChild(root, i)
                Dim typed = TryCast(child, T)
                If typed IsNot Nothing Then
                    Return typed
                End If

                Dim nested = FindVisualDescendant(Of T)(child)
                If nested IsNot Nothing Then
                    Return nested
                End If
            Next

            Return Nothing
        End Function

        Private Shared Sub TrimLogIfNeeded(textBox As TextBox)
            Const maxChars As Integer = 500000
            If textBox.Text.Length <= maxChars Then
                Return
            End If

            textBox.Text = textBox.Text.Substring(textBox.Text.Length - (maxChars \ 2))
            textBox.CaretIndex = textBox.Text.Length
            textBox.ScrollToEnd()
        End Sub
    End Class
End Namespace
