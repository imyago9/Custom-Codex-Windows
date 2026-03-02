Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Text.Json
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
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

        <Flags>
        Private Enum TranscriptScrollRequestReason
            None = 0
            LegacyCallsite = 1
            ThreadSelection = 2
            ThreadRebuild = 4
            RuntimeStream = 8
            SystemMessage = 16
            UserMessage = 32
        End Enum

        Private Enum TranscriptScrollFollowMode
            FollowBottom = 0
            DetachedByUser = 1
        End Enum

        Private Enum TranscriptScrollRequestPolicy
            None = 0
            FollowIfPinned = 1
            ForceJump = 2
        End Enum

        Private Const TranscriptScrollDebugInstrumentationEnabled As Boolean = False

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
            Public Property StagedChangeCount As Integer
            Public Property UnstagedChangeCount As Integer
            Public Property UntrackedChangeCount As Integer
            Public Property ConflictChangeCount As Integer
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

            Public ReadOnly Property IsStagedChange As Boolean
                Get
                    If IsUntracked Then
                        Return False
                    End If

                    Dim code = If(StatusCode, String.Empty).PadRight(2)
                    Dim x = code(0)
                    Return x <> " "c AndAlso x <> "?"c
                End Get
            End Property

            Public ReadOnly Property IsUnstagedChange As Boolean
                Get
                    If IsUntracked Then
                        Return True
                    End If

                    Dim code = If(StatusCode, String.Empty).PadRight(2)
                    Dim y = code(1)
                    Return y <> " "c AndAlso y <> "?"c
                End Get
            End Property

            Public ReadOnly Property StageToggleButtonText As String
                Get
                    If IsUnstagedChange Then
                        Return "Stage"
                    End If

                    If IsStagedChange Then
                        Return "Unstage"
                    End If

                    Return "Stage"
                End Get
            End Property

            Public ReadOnly Property StageToggleToolTip As String
                Get
                    If IsUnstagedChange Then
                        Return "Stage this file"
                    End If

                    If IsStagedChange Then
                        Return "Unstage this file"
                    End If

                    Return "Stage this file"
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
            Public Property OldLineNumber As Integer?
            Public Property NewLineNumber As Integer?
            Public Property SourceIndex As Integer
            Public Property IsInlineEditing As Boolean
            Public Property InlineEditText As String = String.Empty

            Public ReadOnly Property DisplayOldLineNumber As String
                Get
                    Return If(OldLineNumber.HasValue, OldLineNumber.Value.ToString(), String.Empty)
                End Get
            End Property

            Public ReadOnly Property DisplayNewLineNumber As String
                Get
                    Return If(NewLineNumber.HasValue, NewLineNumber.Value.ToString(), String.Empty)
                End Get
            End Property

            Public ReadOnly Property DiffMarker As String
                Get
                    Dim value = If(Text, String.Empty)
                    If value.Length = 0 Then
                        Return " "
                    End If

                    Dim first = value(0)
                    If first = "+"c OrElse first = "-"c OrElse first = " "c Then
                        Return first
                    End If

                    Return " "
                End Get
            End Property

            Public ReadOnly Property DisplayText As String
                Get
                    Dim value = If(Text, String.Empty)
                    If value.Length = 0 Then
                        Return String.Empty
                    End If

                    If StringComparer.Ordinal.Equals(Kind, "added") OrElse
                       StringComparer.Ordinal.Equals(Kind, "removed") OrElse
                       StringComparer.Ordinal.Equals(Kind, "context") Then
                        If value.Length > 0 Then
                            Return value.Substring(1)
                        End If
                    End If

                    Return value
                End Get
            End Property

            Public ReadOnly Property IsEditableLine As Boolean
                Get
                    Return NewLineNumber.HasValue AndAlso
                           (StringComparer.Ordinal.Equals(Kind, "context") OrElse
                            StringComparer.Ordinal.Equals(Kind, "added"))
                End Get
            End Property
        End Class

        Private NotInheritable Class GitInlineDiffEditSession
            Public Property RepoRoot As String = String.Empty
            Public Property RelativePath As String = String.Empty
            Public Property DisplayPath As String = String.Empty
            Public Property FullPath As String = String.Empty
            Public Property FileEncoding As Encoding
            Public Property PreferredNewLine As String = vbLf
            Public Property HadTerminalNewLine As Boolean
            Public Property EditableEntriesByLine As New Dictionary(Of Integer, GitDiffPreviewLineEntry)()
            Public Property OriginalLinesByLine As New Dictionary(Of Integer, String)()
        End Class

        Private Structure EditableTextFileBuffer
            Public Property Lines As List(Of String)
            Public Property FileEncoding As Encoding
            Public Property PreferredNewLine As String
            Public Property HadTerminalNewLine As Boolean
        End Structure

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


        Public NotInheritable Class TurnComposerThreadPreferenceSettings
            Public Property ModelId As String = String.Empty
            Public Property ReasoningEffort As String = "medium"
            Public Property ApprovalPolicy As String = "on-request"
            Public Property Sandbox As String = "workspace-write"
            Public Property CachedContextTurnId As String = String.Empty
            Public Property CachedContextTokenUsageJson As String = String.Empty
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
            Public Property ShowTurnLifecycleDotsInTranscript As Boolean = True
            Public Property ShowReasoningBubblesInTranscript As Boolean = True
            Public Property PlayUiSounds As Boolean = True
            Public Property UiSoundVolumePercent As Double = 100.0R
            Public Property FilterThreadsByWorkingDir As Boolean
            Public Property EncryptedApiKey As String = String.Empty
            Public Property ThemeMode As String = AppAppearanceManager.LightTheme
            Public Property DensityMode As String = AppAppearanceManager.ComfortableDensity
            Public Property TranscriptScaleIndex As Integer = 2
            Public Property TurnComposerPickersCollapsed As Boolean
            Public Property TurnComposerThreadPreferences As Dictionary(Of String, TurnComposerThreadPreferenceSettings) =
                New Dictionary(Of String, TurnComposerThreadPreferenceSettings)(StringComparer.Ordinal)
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
        ' Visible UI selection only. Multi-thread runtime state is tracked separately in the
        ' runtime store/registry and can include active turns in non-visible threads.
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
        Private _logoutUiTransitionInProgress As Boolean
        Private _disconnectUiTransitionInProgress As Boolean
        Private _currentTheme As String = AppAppearanceManager.LightTheme
        Private _currentDensity As String = AppAppearanceManager.ComfortableDensity
        Private _suppressAppearanceUiChange As Boolean
        Private _suppressTranscriptScaleUiChange As Boolean
        Private _turnComposerPickersCollapsed As Boolean
        Private _turnComposerPickersExpandedWidth As Double = 434.0R
        Private _rateLimitIndicatorsIntroPlayed As Boolean
        Private _transcriptScrollFollowMode As TranscriptScrollFollowMode = TranscriptScrollFollowMode.FollowBottom
        Private _suppressTranscriptScrollTracking As Boolean
        Private _transcriptScrollToBottomPending As Boolean
        Private _transcriptScrollToBottomQueuedGeneration As Integer
        Private _transcriptScrollToBottomQueuedPolicy As TranscriptScrollRequestPolicy = TranscriptScrollRequestPolicy.None
        Private _transcriptScrollToBottomQueuedInteractionEpoch As Integer
        Private _transcriptScrollViewer As ScrollViewer
        Private _transcriptScrollThumbDragActive As Boolean
        Private _transcriptUserScrollInteractionArmed As Boolean
        Private _transcriptScrollUserInteractionEpoch As Integer
        Private _transcriptScrollProgrammaticMoveInProgress As Boolean
        Private _transcriptDetachedAnchorOffset As Double?
        Private _transcriptScrollRequestPending As Boolean
        Private _transcriptScrollQueuedPolicy As TranscriptScrollRequestPolicy = TranscriptScrollRequestPolicy.None
        Private _transcriptScrollQueuedReasons As TranscriptScrollRequestReason = TranscriptScrollRequestReason.None
        Private _transcriptScrollQueuedInteractionEpoch As Integer
        Private _transcriptScrollRequestGeneration As Integer
        Private _pendingNewThreadFirstPromptSelection As Boolean
        Private _threadsPanelHintBubbleHintKey As String = String.Empty
        Private _threadsPanelHintBubbleDismissedForCurrentHint As Boolean
        Private _workspaceHintOverlayHintKey As String = String.Empty
        Private _workspaceHintOverlayDismissedForCurrentHint As Boolean
        Private _gitPanelLoadVersion As Integer
        Private _gitPanelDiffPreviewLoadVersion As Integer
        Private _gitPanelCommitPreviewLoadVersion As Integer
        Private _gitPanelBranchPreviewLoadVersion As Integer
        Private _gitPanelLoadingActive As Boolean
        Private _gitPanelCommandInProgress As Boolean
        Private _metricsPanelLoadVersion As Integer
        Private _suppressGitPanelSelectionEvents As Boolean
        Private _gitPanelActiveTab As String = "changes"
        Private _gitPanelDockWidth As Double = 560.0R
        Private _leftSidebarDockWidth As Double = LeftSidebarDefaultDockWidth
        Private _isLeftSidebarVisible As Boolean = True
        Private _currentGitPanelSnapshot As GitPanelSnapshot
        Private _gitPanelSelectedDiffFilePath As String = String.Empty
        Private _gitInlineDiffEditSession As GitInlineDiffEditSession
        Private _gitInlineEditSaveInProgress As Boolean
        Private _gitDiffCtrlDragSelecting As Boolean
        Private _gitDiffCtrlDragAnchorIndex As Integer = -1
        Private _gitDiffCtrlDragLastIndex As Integer = -1
        Private _activeTurnProgressTurnId As String = String.Empty
        Private _activeTurnProgressStartedAtUtc As DateTimeOffset?
        Private _activeTurnProgressWorkingDotsPhase As Integer
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
            _skillsAppsService = New CodexSkillsAppsService(Function() CurrentClient())
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
            AddHandler _viewModel.TranscriptPanel.TranscriptItemsChanged,
                Sub(sender, e)
                    UpdateWorkspaceEmptyStateVisibility()
                End Sub
            AddHandler _viewModel.TranscriptPanel.LeadingEntriesTrimmed,
                AddressOf OnTranscriptLeadingEntriesTrimmed
            AddHandler _viewModel.ThreadsPanel.PropertyChanged,
                Sub(sender, e)
                    If e Is Nothing OrElse
                       Not StringComparer.Ordinal.Equals(e.PropertyName, NameOf(ThreadsPanelViewModel.StateText)) Then
                        Return
                    End If

                    UpdateThreadsPanelStateHintBubbleVisibility()
                End Sub
            AddHandler _viewModel.TurnComposer.PropertyChanged, AddressOf OnTurnComposerPropertyChanged
            InitializeSessionCoordinatorBindings()
            _turnWorkflowCoordinator.BindCommands(AddressOf StartTurnAsync,
                                                  AddressOf SteerTurnAsync,
                                                  AddressOf InterruptTurnAsync,
                                                  AddressOf ResolveApprovalAsync)
            InitializeMvvmCommandBindings()

            InitializeUiDefaults()
            SyncTurnComposerStateForCurrentSelection()
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
            AddHandler SidebarPaneHost.BtnSidebarMetrics.Click, Sub(sender, e) ToggleMetricsPanel()
            AddHandler SidebarPaneHost.BtnSidebarSettings.Click, Sub(sender, e) ShowSettingsView()
            AddHandler SidebarPaneHost.BtnSettingsBack.Click, Sub(sender, e) ShowWorkspaceView()
            AddHandler SidebarPaneHost.CmbDensity.SelectionChanged, Sub(sender, e) OnDensitySelectionChanged()
            AddHandler SidebarPaneHost.CmbTranscriptScale.SelectionChanged, Sub(sender, e) OnTranscriptScaleSelectionChanged()
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
            AddHandler SidebarPaneHost.ChkShowTurnLifecycleDotsInTranscript.Checked,
                Sub(sender, e)
                    SaveSettings()
                    ApplyTranscriptTimelineDotVisibilitySettings()
                End Sub
            AddHandler SidebarPaneHost.ChkShowTurnLifecycleDotsInTranscript.Unchecked,
                Sub(sender, e)
                    SaveSettings()
                    ApplyTranscriptTimelineDotVisibilitySettings()
                End Sub
            AddHandler SidebarPaneHost.ChkShowReasoningBubblesInTranscript.Checked,
                Sub(sender, e)
                    SaveSettings()
                    ApplyTranscriptTimelineDotVisibilitySettings()
                End Sub
            AddHandler SidebarPaneHost.ChkShowReasoningBubblesInTranscript.Unchecked,
                Sub(sender, e)
                    SaveSettings()
                    ApplyTranscriptTimelineDotVisibilitySettings()
                End Sub
            AddHandler SidebarPaneHost.ChkPlayUiSounds.Checked, Sub(sender, e) SaveSettings()
            AddHandler SidebarPaneHost.ChkPlayUiSounds.Unchecked, Sub(sender, e) SaveSettings()
            AddHandler SidebarPaneHost.SldUiSoundVolume.ValueChanged, Sub(sender, e) SaveSettings()

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
            AddHandler WorkspacePaneHost.TxtTurnInput.TextChanged, AddressOf OnTurnInputTextChangedForSuggestions
            AddHandler WorkspacePaneHost.TxtTurnInput.SelectionChanged, AddressOf OnTurnInputSelectionChangedForSuggestions
            AddHandler WorkspacePaneHost.TxtTurnInput.PreviewKeyDown, AddressOf OnTurnInputPreviewKeyDownForSuggestions
            AddHandler WorkspacePaneHost.TxtTurnInput.LostKeyboardFocus, AddressOf OnTurnInputLostKeyboardFocusForSuggestions
            AddHandler WorkspacePaneHost.LstTurnComposerTokenSuggestions.PreviewMouseLeftButtonUp, AddressOf OnTurnComposerTokenSuggestionsPreviewMouseLeftButtonUp
            AddHandler WorkspacePaneHost.LstTurnComposerTokenSuggestions.PreviewKeyDown, AddressOf OnTurnComposerTokenSuggestionsPreviewKeyDown
            AddHandler WorkspacePaneHost.TurnComposerTokenSuggestionsPopup.Closed,
                Sub(sender, e)
                    ClearTurnComposerTokenSuggestionTracking()
                End Sub
            AddHandler GitPaneHost.BtnGitPanelRefresh.Click, Sub(sender, e) FireAndForget(RefreshGitPanelAsync())
            AddHandler GitPaneHost.BtnGitPanelClose.Click, Sub(sender, e) CloseGitPanel()
            AddHandler GitPaneHost.BtnGitTabChanges.Click, Sub(sender, e) ShowGitPanelTab("changes")
            AddHandler GitPaneHost.BtnGitTabHistory.Click, Sub(sender, e) ShowGitPanelTab("history")
            AddHandler GitPaneHost.BtnGitTabBranches.Click, Sub(sender, e) ShowGitPanelTab("branches")
            AddHandler GitPaneHost.BtnGitStageAll.Click, Sub(sender, e) FireAndForget(StageAllGitChangesAsync())
            AddHandler GitPaneHost.BtnGitUnstageAll.Click, Sub(sender, e) FireAndForget(UnstageAllGitChangesAsync())
            AddHandler GitPaneHost.BtnGitPush.Click, Sub(sender, e) FireAndForget(PushGitChangesAsync())
            AddHandler GitPaneHost.BtnGitCommit.Click, Sub(sender, e) FireAndForget(CommitGitChangesAsync())
            AddHandler GitPaneHost.TxtGitCommitMessage.TextChanged, Sub(sender, e) UpdateGitCommitComposerState()
            AddHandler GitPaneHost.ChkGitCommitAmend.Checked, Sub(sender, e) UpdateGitCommitComposerState()
            AddHandler GitPaneHost.ChkGitCommitAmend.Unchecked, Sub(sender, e) UpdateGitCommitComposerState()
            AddHandler GitPaneHost.LstGitPanelChanges.SelectionChanged, AddressOf OnGitPanelChangesSelectionChanged
            AddHandler GitPaneHost.LstGitPanelChanges.MouseDoubleClick, AddressOf OnGitPanelChangesMouseDoubleClick
            GitPaneHost.LstGitPanelChanges.AddHandler(Button.ClickEvent,
                                                            New RoutedEventHandler(AddressOf OnGitPanelChangesListButtonClick),
                                                            True)
            AddHandler GitPaneHost.LstGitPanelCommits.SelectionChanged, AddressOf OnGitPanelCommitsSelectionChanged
            AddHandler GitPaneHost.LstGitPanelBranches.SelectionChanged, AddressOf OnGitPanelBranchesSelectionChanged
            AddHandler GitPaneHost.LstGitPanelDiffPreviewLines.SelectionChanged, AddressOf OnGitPanelDiffPreviewLinesSelectionChanged
            AddHandler GitPaneHost.LstGitPanelDiffPreviewLines.MouseDoubleClick, AddressOf OnGitPanelDiffPreviewLinesMouseDoubleClick
            AddHandler GitPaneHost.LstGitPanelDiffPreviewLines.PreviewMouseLeftButtonDown, AddressOf OnGitPanelDiffPreviewLinesPreviewMouseLeftButtonDown
            AddHandler GitPaneHost.LstGitPanelDiffPreviewLines.PreviewMouseMove, AddressOf OnGitPanelDiffPreviewLinesPreviewMouseMove
            AddHandler GitPaneHost.LstGitPanelDiffPreviewLines.PreviewMouseLeftButtonUp, AddressOf OnGitPanelDiffPreviewLinesPreviewMouseLeftButtonUp
            AddHandler GitPaneHost.LstGitPanelDiffPreviewLines.PreviewKeyDown, AddressOf OnGitPanelDiffPreviewLinesPreviewKeyDown
            AddHandler GitPaneHost.BtnGitDiffInlineStart.Click, Sub(sender, e) BeginGitInlineDiffEditFromSelection()
            AddHandler GitPaneHost.BtnGitDiffInlineSave.Click, Sub(sender, e) FireAndForget(SaveGitInlineDiffEditAsync())
            AddHandler GitPaneHost.BtnGitDiffInlineCancel.Click, Sub(sender, e) CancelGitInlineDiffEdit(shouldShowStatus:=False)
            AddHandler MetricsPaneHost.BtnMetricsPanelRefresh.Click, Sub(sender, e) FireAndForget(RefreshMetricsPanelAsync())
            AddHandler MetricsPaneHost.BtnMetricsPanelClose.Click, Sub(sender, e) CloseMetricsPanel()
            InitializeGitPanelUi()
            InitializeMetricsPanelUi()
            AttachTranscriptInteractionHandlers(WorkspacePaneHost.LstTranscript)
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
            Dim visibilityChanged = (_isLeftSidebarVisible <> isVisible)

            If LeftSidebarColumn Is Nothing Then
                _isLeftSidebarVisible = isVisible
                SyncThreadsSidebarToggleVisual()
                If visibilityChanged Then
                    PlayPanelVisibilityToggleSoundIfEnabled()
                End If
                Return
            End If

            If Not visibilityChanged AndAlso LeftSidebarColumn.ActualWidth >= 1.0R Then
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

            If visibilityChanged Then
                PlayPanelVisibilityToggleSoundIfEnabled()
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

            button.Background = ResolveBrush(If(isSelected, "AccentSubtleBrush", "Transparent"), Brushes.Transparent)
            button.BorderBrush = ResolveBrush(If(isSelected, "ListItemSelectedBorderBrush", "Transparent"), Brushes.Transparent)
            button.BorderThickness = If(isSelected, New Thickness(1), New Thickness(0))
            button.Foreground = ResolveBrush(If(isSelected, "TextPrimaryBrush", "TextSecondaryBrush"), Brushes.Black)
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
            If GitPaneHost.TxtGitCommitMessage IsNot Nothing Then
                GitPaneHost.TxtGitCommitMessage.Text = String.Empty
            End If
            If GitPaneHost.ChkGitCommitAmend IsNot Nothing Then
                GitPaneHost.ChkGitCommitAmend.IsChecked = False
            End If
            If GitPaneHost.LblGitCommitComposerState IsNot Nothing Then
                GitPaneHost.LblGitCommitComposerState.Text = "Load repository status to stage and commit changes."
            End If

            UpdateGitCommitComposerState()
            UpdateGitDiffInlineToolbarState()
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
            If button Is Nothing Then
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
            If StringComparer.Ordinal.Equals(button.Name, "BtnGitChangeOpenInline") Then
                OpenGitChangeInVsCode(selected)
                Return
            End If

            If StringComparer.Ordinal.Equals(button.Name, "BtnGitChangeStageToggle") Then
                FireAndForget(ToggleGitFileStageStateAsync(selected))
                Return
            End If
        End Sub

        Private Sub UpdateGitCommitComposerState(Optional snapshot As GitPanelSnapshot = Nothing)
            If GitPaneHost Is Nothing Then
                Return
            End If

            Dim activeSnapshot = If(snapshot, _currentGitPanelSnapshot)
            Dim hasSnapshot = activeSnapshot IsNot Nothing AndAlso String.IsNullOrWhiteSpace(activeSnapshot.ErrorMessage)
            Dim stagedCount = If(hasSnapshot, activeSnapshot.StagedChangeCount, 0)
            Dim unstagedCount = If(hasSnapshot, activeSnapshot.UnstagedChangeCount, 0)
            Dim untrackedCount = If(hasSnapshot, activeSnapshot.UntrackedChangeCount, 0)
            Dim conflictCount = If(hasSnapshot, activeSnapshot.ConflictChangeCount, 0)
            Dim hasMessage = Not String.IsNullOrWhiteSpace(If(GitPaneHost.TxtGitCommitMessage?.Text, String.Empty))
            Dim isAmendMode = GitPaneHost.ChkGitCommitAmend IsNot Nothing AndAlso
                              GitPaneHost.ChkGitCommitAmend.IsChecked.HasValue AndAlso
                              GitPaneHost.ChkGitCommitAmend.IsChecked.Value
            Dim branchName = If(If(activeSnapshot?.BranchName, String.Empty), String.Empty).Trim()
            Dim canPushBranch = hasSnapshot AndAlso
                                Not String.IsNullOrWhiteSpace(branchName) AndAlso
                                Not StringComparer.OrdinalIgnoreCase.Equals(branchName, "HEAD")
            Dim hasAnyLocalChanges = stagedCount > 0 OrElse
                                     unstagedCount > 0 OrElse
                                     untrackedCount > 0 OrElse
                                     conflictCount > 0
            Dim showPushPrimary = hasSnapshot AndAlso canPushBranch AndAlso Not hasAnyLocalChanges
            Dim allowInteraction = hasSnapshot AndAlso Not _gitPanelLoadingActive AndAlso Not _gitPanelCommandInProgress

            If GitPaneHost.BtnGitStageAll IsNot Nothing Then
                GitPaneHost.BtnGitStageAll.IsEnabled = allowInteraction AndAlso (unstagedCount > 0 OrElse untrackedCount > 0)
            End If

            If GitPaneHost.BtnGitUnstageAll IsNot Nothing Then
                GitPaneHost.BtnGitUnstageAll.IsEnabled = allowInteraction AndAlso stagedCount > 0
            End If

            Dim canCommit As Boolean
            If isAmendMode Then
                canCommit = allowInteraction AndAlso (stagedCount > 0 OrElse hasMessage)
            Else
                canCommit = allowInteraction AndAlso stagedCount > 0 AndAlso hasMessage
            End If

            If GitPaneHost.BtnGitPush IsNot Nothing Then
                GitPaneHost.BtnGitPush.Visibility = If(showPushPrimary, Visibility.Visible, Visibility.Collapsed)
                GitPaneHost.BtnGitPush.IsEnabled = allowInteraction AndAlso showPushPrimary
            End If

            If GitPaneHost.BtnGitCommit IsNot Nothing Then
                GitPaneHost.BtnGitCommit.Visibility = If(showPushPrimary, Visibility.Collapsed, Visibility.Visible)
                GitPaneHost.BtnGitCommit.IsEnabled = canCommit
            End If

            If GitPaneHost.ChkGitCommitAmend IsNot Nothing Then
                GitPaneHost.ChkGitCommitAmend.IsEnabled = hasSnapshot AndAlso Not _gitPanelLoadingActive AndAlso Not _gitPanelCommandInProgress
            End If

            If GitPaneHost.TxtGitCommitMessage IsNot Nothing Then
                GitPaneHost.TxtGitCommitMessage.IsEnabled = hasSnapshot AndAlso Not _gitPanelCommandInProgress
            End If

            If GitPaneHost.LblGitCommitComposerState Is Nothing Then
                Return
            End If

            If _gitPanelCommandInProgress Then
                GitPaneHost.LblGitCommitComposerState.Text = "Running git command..."
                Return
            End If

            If _gitPanelLoadingActive Then
                GitPaneHost.LblGitCommitComposerState.Text = "Loading repository status..."
                Return
            End If

            If Not hasSnapshot Then
                GitPaneHost.LblGitCommitComposerState.Text = "Load repository status to stage and commit changes."
                Return
            End If

            Dim statusParts As New List(Of String)()
            If stagedCount > 0 Then
                statusParts.Add($"{stagedCount} staged")
            End If
            If unstagedCount > 0 Then
                statusParts.Add($"{unstagedCount} unstaged")
            End If
            If untrackedCount > 0 Then
                statusParts.Add($"{untrackedCount} untracked")
            End If
            If conflictCount > 0 Then
                statusParts.Add($"{conflictCount} conflicts")
            End If

            If statusParts.Count = 0 Then
                If isAmendMode Then
                    GitPaneHost.LblGitCommitComposerState.Text = If(hasMessage,
                                                                    "Working tree clean. Amend is on; commit will amend message.",
                                                                    "Working tree clean. Amend is on; add a message to amend last commit.")
                ElseIf showPushPrimary Then
                    GitPaneHost.LblGitCommitComposerState.Text = "Working tree clean. Ready to push."
                Else
                    GitPaneHost.LblGitCommitComposerState.Text = "Working tree clean."
                End If
                Return
            End If

            Dim statusSummary = String.Join(", ", statusParts)
            If isAmendMode Then
                If stagedCount > 0 AndAlso hasMessage Then
                    GitPaneHost.LblGitCommitComposerState.Text = $"{statusSummary}. Amend is on; commit will amend with new message."
                ElseIf stagedCount > 0 Then
                    GitPaneHost.LblGitCommitComposerState.Text = $"{statusSummary}. Amend is on; commit will amend without changing message."
                ElseIf hasMessage Then
                    GitPaneHost.LblGitCommitComposerState.Text = $"{statusSummary}. Amend is on; commit will amend message only."
                Else
                    GitPaneHost.LblGitCommitComposerState.Text = $"{statusSummary}. Amend is on; add a message or stage changes."
                End If
            Else
                If stagedCount > 0 Then
                    GitPaneHost.LblGitCommitComposerState.Text = If(hasMessage,
                                                                    $"{statusSummary}. Ready to commit.",
                                                                    $"{statusSummary}. Add a commit message.")
                Else
                    GitPaneHost.LblGitCommitComposerState.Text = $"{statusSummary}. Stage changes to enable commit."
                End If
            End If

            If Not canPushBranch AndAlso StringComparer.OrdinalIgnoreCase.Equals(branchName, "HEAD") Then
                GitPaneHost.LblGitCommitComposerState.Text &= " Push unavailable on detached HEAD."
            End If
        End Sub

        Private Sub SetGitPanelCommandBusyState(isBusy As Boolean, Optional statusText As String = Nothing)
            _gitPanelCommandInProgress = isBusy

            If Not String.IsNullOrWhiteSpace(statusText) AndAlso GitPaneHost.LblGitPanelState IsNot Nothing Then
                GitPaneHost.LblGitPanelState.Text = statusText
            End If

            If GitPaneHost.BtnGitPanelRefresh IsNot Nothing Then
                GitPaneHost.BtnGitPanelRefresh.IsEnabled = Not _gitPanelLoadingActive AndAlso Not _gitPanelCommandInProgress
            End If
            If GitPaneHost.LstGitPanelChanges IsNot Nothing Then
                GitPaneHost.LstGitPanelChanges.IsEnabled = Not _gitPanelCommandInProgress
            End If

            UpdateGitCommitComposerState()
            UpdateGitDiffInlineToolbarState()
        End Sub

        Private Async Function ToggleGitFileStageStateAsync(selected As GitChangedFileListEntry) As Task
            If selected Is Nothing OrElse _gitPanelCommandInProgress Then
                Return
            End If

            Dim repoRoot = ResolveCurrentGitPanelRepoRoot()
            If String.IsNullOrWhiteSpace(repoRoot) Then
                ShowStatus("Git panel repository is not available.", isError:=True, displayToast:=True)
                Return
            End If

            Dim filePath = If(selected.FilePath, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(filePath) Then
                ShowStatus("No file path available for this entry.", isError:=True, displayToast:=True)
                Return
            End If

            Dim shouldStage = selected.IsUnstagedChange OrElse selected.IsUntracked OrElse Not selected.IsStagedChange
            Dim actionLabel = If(shouldStage, "Staging", "Unstaging")
            Dim quotedPath = QuoteProcessArgument(filePath)
            Dim commandArgs = If(shouldStage,
                                 $"add -- {quotedPath}",
                                 $"restore --staged -- {quotedPath}")

            SetGitPanelCommandBusyState(True, $"{actionLabel} {selected.DisplayPath}...")
            Try
                Dim result = Await Task.Run(Function() RunProcessCapture("git", commandArgs, repoRoot)).ConfigureAwait(True)
                If Not shouldStage AndAlso result.ExitCode <> 0 Then
                    result = Await Task.Run(Function() RunProcessCapture("git", $"reset HEAD -- {quotedPath}", repoRoot)).ConfigureAwait(True)
                End If

                If result.ExitCode <> 0 Then
                    Dim reason = FirstNonEmptyLine(NormalizeProcessError(result))
                    ShowStatus($"Could not {If(shouldStage, "stage", "unstage")} {selected.DisplayPath}: {If(String.IsNullOrWhiteSpace(reason), "unknown error", reason)}",
                               isError:=True,
                               displayToast:=True)
                    Return
                End If

                ShowStatus($"{If(shouldStage, "Staged", "Unstaged")} {selected.DisplayPath}")
                Await RefreshGitPanelAsync().ConfigureAwait(True)
            Finally
                SetGitPanelCommandBusyState(False)
            End Try
        End Function

        Private Async Function StageAllGitChangesAsync() As Task
            If _gitPanelCommandInProgress Then
                Return
            End If

            Dim snapshot = _currentGitPanelSnapshot
            If snapshot Is Nothing Then
                ShowStatus("No repository snapshot loaded.", isError:=True, displayToast:=True)
                Return
            End If
            If String.IsNullOrWhiteSpace(snapshot.RepoRoot) Then
                ShowStatus("Git panel repository is not available.", isError:=True, displayToast:=True)
                Return
            End If

            If snapshot.UnstagedChangeCount <= 0 AndAlso snapshot.UntrackedChangeCount <= 0 Then
                ShowStatus("No unstaged or untracked changes to stage.")
                Return
            End If

            SetGitPanelCommandBusyState(True, "Staging all changes...")
            Try
                Dim result = Await Task.Run(Function() RunProcessCapture("git", "add -A", snapshot.RepoRoot)).ConfigureAwait(True)
                If result.ExitCode <> 0 Then
                    Dim reason = FirstNonEmptyLine(NormalizeProcessError(result))
                    ShowStatus($"Could not stage all changes: {If(String.IsNullOrWhiteSpace(reason), "unknown error", reason)}", isError:=True, displayToast:=True)
                    Return
                End If

                ShowStatus("Staged all changes.")
                Await RefreshGitPanelAsync().ConfigureAwait(True)
            Finally
                SetGitPanelCommandBusyState(False)
            End Try
        End Function

        Private Async Function UnstageAllGitChangesAsync() As Task
            If _gitPanelCommandInProgress Then
                Return
            End If

            Dim snapshot = _currentGitPanelSnapshot
            If snapshot Is Nothing Then
                ShowStatus("No repository snapshot loaded.", isError:=True, displayToast:=True)
                Return
            End If
            If String.IsNullOrWhiteSpace(snapshot.RepoRoot) Then
                ShowStatus("Git panel repository is not available.", isError:=True, displayToast:=True)
                Return
            End If

            If snapshot.StagedChangeCount <= 0 Then
                ShowStatus("No staged changes to unstage.")
                Return
            End If

            SetGitPanelCommandBusyState(True, "Unstaging all changes...")
            Try
                Dim result = Await Task.Run(Function() RunProcessCapture("git", "reset", snapshot.RepoRoot)).ConfigureAwait(True)
                If result.ExitCode <> 0 Then
                    Dim reason = FirstNonEmptyLine(NormalizeProcessError(result))
                    ShowStatus($"Could not unstage all changes: {If(String.IsNullOrWhiteSpace(reason), "unknown error", reason)}", isError:=True, displayToast:=True)
                    Return
                End If

                ShowStatus("Unstaged all changes.")
                Await RefreshGitPanelAsync().ConfigureAwait(True)
            Finally
                SetGitPanelCommandBusyState(False)
            End Try
        End Function

        Private Async Function PushGitChangesAsync() As Task
            If _gitPanelCommandInProgress Then
                Return
            End If

            Dim snapshot = _currentGitPanelSnapshot
            If snapshot Is Nothing OrElse String.IsNullOrWhiteSpace(snapshot.RepoRoot) Then
                ShowStatus("Git panel repository is not available.", isError:=True, displayToast:=True)
                Return
            End If

            Dim branchName = If(snapshot.BranchName, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(branchName) OrElse StringComparer.OrdinalIgnoreCase.Equals(branchName, "HEAD") Then
                Dim branchResult = Await Task.Run(Function() RunProcessCapture("git", "rev-parse --abbrev-ref HEAD", snapshot.RepoRoot)).ConfigureAwait(True)
                If branchResult.ExitCode = 0 Then
                    branchName = FirstNonEmptyLine(branchResult.OutputText)
                End If
            End If

            branchName = If(branchName, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(branchName) OrElse StringComparer.OrdinalIgnoreCase.Equals(branchName, "HEAD") Then
                ShowStatus("Cannot push while on a detached HEAD.", isError:=True, displayToast:=True)
                Return
            End If

            SetGitPanelCommandBusyState(True, $"Pushing {branchName}...")
            Try
                Dim upstreamResult = Await Task.Run(
                    Function()
                        Return RunProcessCapture("git",
                                                 "rev-parse --abbrev-ref --symbolic-full-name @{u}",
                                                 snapshot.RepoRoot)
                    End Function).ConfigureAwait(True)

                Dim hasUpstream = upstreamResult.ExitCode = 0 AndAlso
                                  Not String.IsNullOrWhiteSpace(FirstNonEmptyLine(upstreamResult.OutputText))
                Dim pushArgs = If(hasUpstream,
                                  "push",
                                  $"push -u origin {QuoteProcessArgument(branchName)}")

                Dim pushResult = Await Task.Run(Function() RunProcessCapture("git", pushArgs, snapshot.RepoRoot)).ConfigureAwait(True)
                If pushResult.ExitCode <> 0 Then
                    Dim reason = FirstNonEmptyLine(NormalizeProcessError(pushResult))
                    ShowStatus($"Push failed: {If(String.IsNullOrWhiteSpace(reason), "unknown error", reason)}",
                               isError:=True,
                               displayToast:=True)
                    Return
                End If

                ShowStatus(ResolveGitCommandSummary(pushResult, $"Pushed {branchName}."), displayToast:=True)
                PlayGitPushSoundIfEnabled()
                Await RefreshGitPanelAsync().ConfigureAwait(True)
            Finally
                SetGitPanelCommandBusyState(False)
            End Try
        End Function

        Private Async Function CommitGitChangesAsync() As Task
            If _gitPanelCommandInProgress Then
                Return
            End If

            Dim snapshot = _currentGitPanelSnapshot
            If snapshot Is Nothing OrElse String.IsNullOrWhiteSpace(snapshot.RepoRoot) Then
                ShowStatus("Git panel repository is not available.", isError:=True, displayToast:=True)
                Return
            End If

            Dim commitMessage = If(GitPaneHost.TxtGitCommitMessage?.Text, String.Empty)
            Dim hasMessage = Not String.IsNullOrWhiteSpace(commitMessage)
            Dim isAmendMode = GitPaneHost.ChkGitCommitAmend IsNot Nothing AndAlso
                              GitPaneHost.ChkGitCommitAmend.IsChecked.HasValue AndAlso
                              GitPaneHost.ChkGitCommitAmend.IsChecked.Value

            If Not isAmendMode Then
                If Not hasMessage Then
                    ShowStatus("Enter a commit message first.", isError:=True, displayToast:=True)
                    Return
                End If

                If snapshot.StagedChangeCount <= 0 Then
                    ShowStatus("Stage at least one file before committing.", isError:=True, displayToast:=True)
                    Return
                End If
            Else
                If snapshot.StagedChangeCount <= 0 AndAlso Not hasMessage Then
                    ShowStatus("For amend, stage changes or provide a new message.", isError:=True, displayToast:=True)
                    Return
                End If
            End If

            Dim commitMessageFilePath As String = String.Empty

            SetGitPanelCommandBusyState(True, If(isAmendMode, "Amending commit...", "Creating commit..."))
            Try
                If hasMessage Then
                    commitMessageFilePath = Path.Combine(Path.GetTempPath(), $"codex-git-commit-{Guid.NewGuid():N}.txt")
                    Dim commitBody = commitMessage.TrimEnd() & Environment.NewLine
                    File.WriteAllText(commitMessageFilePath, commitBody, New UTF8Encoding(False))
                End If

                Dim commitArgs As String
                If isAmendMode Then
                    If hasMessage Then
                        commitArgs = $"commit --amend -F {QuoteProcessArgument(commitMessageFilePath)}"
                    Else
                        commitArgs = "commit --amend --no-edit"
                    End If
                Else
                    commitArgs = $"commit -F {QuoteProcessArgument(commitMessageFilePath)}"
                End If

                Dim result = Await Task.Run(Function() RunProcessCapture("git", commitArgs, snapshot.RepoRoot)).ConfigureAwait(True)

                If result.ExitCode <> 0 Then
                    Dim reason = FirstNonEmptyLine(NormalizeProcessError(result))
                    ShowStatus($"Commit failed: {If(String.IsNullOrWhiteSpace(reason), "unknown error", reason)}",
                               isError:=True,
                               displayToast:=True)
                    Return
                End If

                If hasMessage AndAlso GitPaneHost.TxtGitCommitMessage IsNot Nothing Then
                    GitPaneHost.TxtGitCommitMessage.Text = String.Empty
                End If

                ShowStatus(ResolveGitCommandSummary(result, If(isAmendMode, "Commit amended.", "Commit created.")), displayToast:=True)
                PlayGitCommitSoundIfEnabled()
                Await RefreshGitPanelAsync().ConfigureAwait(True)
            Finally
                Try
                    If Not String.IsNullOrWhiteSpace(commitMessageFilePath) AndAlso File.Exists(commitMessageFilePath) Then
                        File.Delete(commitMessageFilePath)
                    End If
                Catch
                End Try

                SetGitPanelCommandBusyState(False)
            End Try
        End Function

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

            EndGitDiffCtrlDragSelection()
            CancelGitInlineDiffEdit(shouldShowStatus:=False)

            Dim lines = BuildGitDiffPreviewLineEntries(previewText)
            GitPaneHost.LstGitPanelDiffPreviewLines.ItemsSource = Nothing
            GitPaneHost.LstGitPanelDiffPreviewLines.ItemsSource = lines

            If lines IsNot Nothing AndAlso lines.Count > 0 Then
                GitPaneHost.LstGitPanelDiffPreviewLines.ScrollIntoView(lines(0))
            End If

            UpdateGitDiffInlineToolbarState()
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
            Dim currentOldLineNumber As Integer = 0
            Dim currentNewLineNumber As Integer = 0
            Dim inHunk As Boolean
            Dim sourceIndex As Integer = 0

            For Each line In lines
                If ShouldHideGitDiffHeaderLineFromPreview(line) Then
                    Continue For
                End If

                Dim kind = ClassifyGitDiffPreviewLine(line)
                Dim entry As New GitDiffPreviewLineEntry() With {
                    .Text = line,
                    .Kind = kind,
                    .SourceIndex = sourceIndex
                }

                If kind = "hunk" Then
                    Dim parsedOldStart As Integer
                    Dim parsedOldCount As Integer
                    Dim parsedNewStart As Integer
                    Dim parsedNewCount As Integer
                    If ParseGitDiffHunkHeader(line, parsedOldStart, parsedOldCount, parsedNewStart, parsedNewCount) Then
                        currentOldLineNumber = parsedOldStart
                        currentNewLineNumber = parsedNewStart
                        inHunk = True
                    Else
                        inHunk = False
                    End If
                ElseIf inHunk Then
                    If line.StartsWith(" ", StringComparison.Ordinal) Then
                        entry.OldLineNumber = currentOldLineNumber
                        entry.NewLineNumber = currentNewLineNumber
                        currentOldLineNumber += 1
                        currentNewLineNumber += 1
                    ElseIf line.StartsWith("+", StringComparison.Ordinal) AndAlso Not line.StartsWith("+++", StringComparison.Ordinal) Then
                        entry.NewLineNumber = currentNewLineNumber
                        currentNewLineNumber += 1
                    ElseIf line.StartsWith("-", StringComparison.Ordinal) AndAlso Not line.StartsWith("---", StringComparison.Ordinal) Then
                        entry.OldLineNumber = currentOldLineNumber
                        currentOldLineNumber += 1
                    ElseIf line.StartsWith("\", StringComparison.Ordinal) Then
                    ElseIf kind = "fileHeader" OrElse kind = "pathHeader" OrElse kind = "meta" Then
                        inHunk = False
                    End If
                End If

                result.Add(entry)
                sourceIndex += 1
            Next

            Return result
        End Function

        Private Shared Function ShouldHideGitDiffHeaderLineFromPreview(line As String) As Boolean
            Dim value = If(line, String.Empty)
            Return value.StartsWith("diff --git ", StringComparison.Ordinal) OrElse
                   value.StartsWith("index ", StringComparison.Ordinal) OrElse
                   value.StartsWith("--- ", StringComparison.Ordinal) OrElse
                   value.StartsWith("+++ ", StringComparison.Ordinal)
        End Function

        Private Shared Function ParseGitDiffHunkHeader(line As String,
                                                       ByRef oldStart As Integer,
                                                       ByRef oldCount As Integer,
                                                       ByRef newStart As Integer,
                                                       ByRef newCount As Integer) As Boolean
            oldStart = 0
            oldCount = 0
            newStart = 0
            newCount = 0

            Dim value = If(line, String.Empty)
            If String.IsNullOrWhiteSpace(value) Then
                Return False
            End If

            Dim match = Regex.Match(value, "^@@\s*-(\d+)(?:,(\d+))?\s+\+(\d+)(?:,(\d+))?\s+@@")
            If Not match.Success Then
                Return False
            End If

            If Not Integer.TryParse(match.Groups(1).Value, oldStart) Then
                Return False
            End If
            If match.Groups(2).Success AndAlso Not Integer.TryParse(match.Groups(2).Value, oldCount) Then
                Return False
            End If
            If Not match.Groups(2).Success Then
                oldCount = 1
            End If

            If Not Integer.TryParse(match.Groups(3).Value, newStart) Then
                Return False
            End If
            If match.Groups(4).Success AndAlso Not Integer.TryParse(match.Groups(4).Value, newCount) Then
                Return False
            End If
            If Not match.Groups(4).Success Then
                newCount = 1
            End If

            Return True
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

        Private Sub OnGitPanelDiffPreviewLinesSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If _suppressGitPanelSelectionEvents Then
                Return
            End If

            UpdateGitDiffInlineToolbarState()
        End Sub

        Private Sub OnGitPanelDiffPreviewLinesMouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
            If e Is Nothing OrElse _gitInlineEditSaveInProgress Then
                Return
            End If

            Dim source = TryCast(e.OriginalSource, DependencyObject)
            If source Is Nothing Then
                Return
            End If

            If FindVisualAncestor(Of ScrollBar)(source) IsNot Nothing Then
                Return
            End If

            Dim listItem = FindVisualAncestor(Of ListBoxItem)(source)
            If listItem Is Nothing Then
                Return
            End If

            Dim clickedLine = TryCast(listItem.DataContext, GitDiffPreviewLineEntry)
            BeginGitInlineDiffEditFromSelection(clickedLine)
            e.Handled = True
        End Sub

        Private Sub OnGitPanelDiffPreviewLinesPreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
            If e Is Nothing Then
                Return
            End If

            If _gitInlineDiffEditSession IsNot Nothing Then
                Return
            End If

            If (Keyboard.Modifiers And ModifierKeys.Control) <> ModifierKeys.Control Then
                EndGitDiffCtrlDragSelection()
                Return
            End If

            Dim source = TryCast(e.OriginalSource, DependencyObject)
            If source Is Nothing Then
                Return
            End If

            If FindVisualAncestor(Of ScrollBar)(source) IsNot Nothing Then
                Return
            End If

            Dim listItem = FindVisualAncestor(Of ListBoxItem)(source)
            If listItem Is Nothing Then
                EndGitDiffCtrlDragSelection()
                Return
            End If

            Dim list = GitPaneHost.LstGitPanelDiffPreviewLines
            If list Is Nothing Then
                EndGitDiffCtrlDragSelection()
                Return
            End If

            list.Focus()
            Dim anchorIndex = ResolveGitDiffListItemIndex(listItem)
            If anchorIndex < 0 Then
                EndGitDiffCtrlDragSelection()
                Return
            End If
            EnsureGitDiffLineSelection(anchorIndex, anchorIndex)

            Dim clickedLine = TryCast(listItem.DataContext, GitDiffPreviewLineEntry)
            If e.ClickCount >= 2 Then
                EndGitDiffCtrlDragSelection()
                BeginGitInlineDiffEditFromSelection(clickedLine)
                e.Handled = True
                Return
            End If

            _gitDiffCtrlDragSelecting = True
            _gitDiffCtrlDragAnchorIndex = anchorIndex
            _gitDiffCtrlDragLastIndex = anchorIndex
            Mouse.Capture(list)
            e.Handled = True
        End Sub

        Private Sub OnGitPanelDiffPreviewLinesPreviewMouseMove(sender As Object, e As MouseEventArgs)
            If e Is Nothing OrElse Not _gitDiffCtrlDragSelecting Then
                Return
            End If

            Dim list = GitPaneHost?.LstGitPanelDiffPreviewLines
            If list Is Nothing Then
                EndGitDiffCtrlDragSelection()
                Return
            End If

            If e.LeftButton <> MouseButtonState.Pressed OrElse
               (Keyboard.Modifiers And ModifierKeys.Control) <> ModifierKeys.Control Then
                EndGitDiffCtrlDragSelection()
                Return
            End If

            Dim point = e.GetPosition(list)
            If point.X < 0 OrElse point.Y < 0 OrElse point.X > list.ActualWidth OrElse point.Y > list.ActualHeight Then
                Return
            End If

            Dim hoverIndex = ResolveGitDiffListItemIndexAtPoint(list, point)
            If hoverIndex < 0 Then
                Return
            End If

            If _gitDiffCtrlDragAnchorIndex < 0 Then
                _gitDiffCtrlDragAnchorIndex = hoverIndex
                _gitDiffCtrlDragLastIndex = hoverIndex
                EnsureGitDiffLineSelection(_gitDiffCtrlDragAnchorIndex, hoverIndex)
                e.Handled = True
                Return
            End If

            If hoverIndex <> _gitDiffCtrlDragLastIndex Then
                EnsureGitDiffLineSelection(_gitDiffCtrlDragAnchorIndex, hoverIndex)
                _gitDiffCtrlDragLastIndex = hoverIndex
            End If

            e.Handled = True
        End Sub

        Private Sub OnGitPanelDiffPreviewLinesPreviewMouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
            If Not _gitDiffCtrlDragSelecting Then
                Return
            End If

            EndGitDiffCtrlDragSelection()
            If e IsNot Nothing Then
                e.Handled = True
            End If
        End Sub

        Private Sub OnGitPanelDiffPreviewLinesPreviewKeyDown(sender As Object, e As KeyEventArgs)
            If e Is Nothing Then
                Return
            End If

            Dim hasControl = (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control
            If _gitInlineDiffEditSession IsNot Nothing Then
                If e.Key = Key.Escape Then
                    CancelGitInlineDiffEdit(shouldShowStatus:=False)
                    e.Handled = True
                    Return
                End If

                If (hasControl AndAlso (e.Key = Key.S OrElse e.Key = Key.Enter)) OrElse
                   (Not hasControl AndAlso e.Key = Key.Enter) Then
                    FireAndForget(SaveGitInlineDiffEditAsync())
                    e.Handled = True
                    Return
                End If

                Return
            End If

            If e.Key = Key.F2 OrElse (hasControl AndAlso e.Key = Key.E) Then
                BeginGitInlineDiffEditFromSelection()
                e.Handled = True
            End If
        End Sub

        Private Sub BeginGitInlineDiffEditFromSelection(Optional clickedLine As GitDiffPreviewLineEntry = Nothing)
            If GitPaneHost Is Nothing OrElse
               GitPaneHost.GitDiffInlineEditActions Is Nothing Then
                Return
            End If

            If _gitInlineEditSaveInProgress Then
                Return
            End If

            If _gitInlineDiffEditSession IsNot Nothing Then
                CancelGitInlineDiffEdit(shouldShowStatus:=False)
            End If

            Dim session As GitInlineDiffEditSession = Nothing
            Dim reason As String = String.Empty
            If Not TryBuildGitInlineDiffEditSession(session, reason, clickedLine) Then
                If Not String.IsNullOrWhiteSpace(reason) Then
                    ShowStatus(reason, isError:=True, displayToast:=True)
                End If
                Return
            End If

            _gitInlineDiffEditSession = session
            RefreshGitDiffPreviewListBindings()
            UpdateGitDiffInlineToolbarState()
            FocusFirstGitInlineEditor()
        End Sub

        Private Function ResolveSelectedEditableGitDiffEntries(Optional clickedLine As GitDiffPreviewLineEntry = Nothing) As List(Of GitDiffPreviewLineEntry)
            Dim result As New List(Of GitDiffPreviewLineEntry)()
            Dim seenSourceIndexes As New HashSet(Of Integer)()
            Dim list = GitPaneHost?.LstGitPanelDiffPreviewLines
            If list Is Nothing Then
                Return result
            End If

            For Each item In list.SelectedItems
                Dim entry = TryCast(item, GitDiffPreviewLineEntry)
                If entry Is Nothing OrElse Not entry.IsEditableLine Then
                    Continue For
                End If

                If seenSourceIndexes.Add(entry.SourceIndex) Then
                    result.Add(entry)
                End If
            Next

            If result.Count = 0 Then
                Dim selectedEntry = TryCast(list.SelectedItem, GitDiffPreviewLineEntry)
                If selectedEntry IsNot Nothing AndAlso selectedEntry.IsEditableLine Then
                    If seenSourceIndexes.Add(selectedEntry.SourceIndex) Then
                        result.Add(selectedEntry)
                    End If
                End If
            End If

            If result.Count = 0 AndAlso clickedLine IsNot Nothing AndAlso clickedLine.IsEditableLine Then
                If seenSourceIndexes.Add(clickedLine.SourceIndex) Then
                    result.Add(clickedLine)
                End If
            End If

            If result.Count = 0 AndAlso clickedLine IsNot Nothing Then
                Dim entries = TryCast(list.ItemsSource, List(Of GitDiffPreviewLineEntry))
                If entries IsNot Nothing AndAlso entries.Count > 0 Then
                    Dim anchorIndex = clickedLine.SourceIndex
                    If anchorIndex < 0 OrElse anchorIndex >= entries.Count Then
                        anchorIndex = entries.IndexOf(clickedLine)
                    End If

                    If anchorIndex >= 0 AndAlso anchorIndex < entries.Count Then
                        For radius = 1 To entries.Count
                            Dim beforeIndex = anchorIndex - radius
                            If beforeIndex >= 0 Then
                                Dim beforeLine = entries(beforeIndex)
                                If beforeLine IsNot Nothing AndAlso beforeLine.IsEditableLine Then
                                    If seenSourceIndexes.Add(beforeLine.SourceIndex) Then
                                        result.Add(beforeLine)
                                    End If
                                    Exit For
                                End If
                            End If

                            Dim afterIndex = anchorIndex + radius
                            If afterIndex < entries.Count Then
                                Dim afterLine = entries(afterIndex)
                                If afterLine IsNot Nothing AndAlso afterLine.IsEditableLine Then
                                    If seenSourceIndexes.Add(afterLine.SourceIndex) Then
                                        result.Add(afterLine)
                                    End If
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End If
            End If

            result.Sort(Function(a, b) a.SourceIndex.CompareTo(b.SourceIndex))
            Return result
        End Function

        Private Function TryBuildGitInlineDiffEditSession(ByRef session As GitInlineDiffEditSession,
                                                          ByRef failureReason As String,
                                                          Optional clickedLine As GitDiffPreviewLineEntry = Nothing) As Boolean
            session = Nothing
            failureReason = String.Empty

            If GitPaneHost Is Nothing OrElse
               GitPaneHost.LstGitPanelDiffPreviewLines Is Nothing OrElse
               GitPaneHost.LstGitPanelChanges Is Nothing Then
                failureReason = "Diff editor is unavailable right now."
                Return False
            End If

            Dim selectedFile = TryCast(GitPaneHost.LstGitPanelChanges.SelectedItem, GitChangedFileListEntry)
            If selectedFile Is Nothing Then
                failureReason = "Select a changed file first."
                Return False
            End If

            Dim repoRoot = ResolveCurrentGitPanelRepoRoot()
            If String.IsNullOrWhiteSpace(repoRoot) Then
                failureReason = "Git repository is unavailable."
                Return False
            End If

            Dim relativePath = If(selectedFile.FilePath, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(relativePath) Then
                failureReason = "No file path is available for this diff."
                Return False
            End If

            Dim selectedEntries = ResolveSelectedEditableGitDiffEntries(clickedLine)
            If selectedEntries.Count = 0 Then
                failureReason = "Select one or more added/context lines, then start editing."
                Return False
            End If

            Dim repoRootFull = Path.GetFullPath(repoRoot).TrimEnd("\"c, "/"c)
            Dim fullPath = Path.GetFullPath(Path.Combine(repoRootFull, relativePath))
            Dim repoRootWithSlash = repoRootFull & Path.DirectorySeparatorChar
            If Not fullPath.StartsWith(repoRootWithSlash, StringComparison.OrdinalIgnoreCase) AndAlso
               Not StringComparer.OrdinalIgnoreCase.Equals(fullPath, repoRootFull) Then
                failureReason = "Selected file is outside the repository root."
                Return False
            End If

            If Not File.Exists(fullPath) Then
                failureReason = "The selected file does not exist on disk."
                Return False
            End If

            Dim fileBuffer As EditableTextFileBuffer
            Try
                fileBuffer = ReadEditableTextFileBuffer(fullPath)
            Catch ex As Exception
                failureReason = $"Could not read file for inline editing: {ex.Message}"
                Return False
            End Try

            If fileBuffer.Lines Is Nothing OrElse fileBuffer.Lines.Count = 0 Then
                failureReason = "The selected file has no editable text lines."
                Return False
            End If

            session = New GitInlineDiffEditSession() With {
                .RepoRoot = repoRootFull,
                .RelativePath = relativePath,
                .DisplayPath = If(selectedFile.DisplayPath, relativePath),
                .FullPath = fullPath,
                .FileEncoding = fileBuffer.FileEncoding,
                .PreferredNewLine = If(String.IsNullOrEmpty(fileBuffer.PreferredNewLine), vbLf, fileBuffer.PreferredNewLine),
                .HadTerminalNewLine = fileBuffer.HadTerminalNewLine
            }

            selectedEntries.Sort(Function(a, b) a.SourceIndex.CompareTo(b.SourceIndex))
            For Each entry In selectedEntries
                If entry Is Nothing OrElse Not entry.IsEditableLine OrElse Not entry.NewLineNumber.HasValue Then
                    Continue For
                End If

                Dim lineNumber = entry.NewLineNumber.Value
                If lineNumber <= 0 OrElse lineNumber > fileBuffer.Lines.Count Then
                    failureReason = "Selected lines are out of date. Refresh diff and try again."
                    Return False
                End If

                If session.EditableEntriesByLine.ContainsKey(lineNumber) Then
                    Continue For
                End If

                Dim originalLine = fileBuffer.Lines(lineNumber - 1)
                session.EditableEntriesByLine(lineNumber) = entry
                session.OriginalLinesByLine(lineNumber) = originalLine
            Next

            If session.EditableEntriesByLine.Count = 0 Then
                failureReason = "No editable diff lines were selected."
                Return False
            End If

            For Each pair In session.EditableEntriesByLine
                Dim lineNumber = pair.Key
                Dim entry = pair.Value
                entry.IsInlineEditing = True
                entry.InlineEditText = session.OriginalLinesByLine(lineNumber)
            Next

            Return True
        End Function

        Private Async Function SaveGitInlineDiffEditAsync() As Task
            Dim session = _gitInlineDiffEditSession
            If session Is Nothing OrElse _gitInlineEditSaveInProgress Then
                Return
            End If

            If GitPaneHost Is Nothing OrElse
               GitPaneHost.BtnGitDiffInlineSave Is Nothing OrElse
               GitPaneHost.BtnGitDiffInlineCancel Is Nothing Then
                Return
            End If

            Dim hasAnyChange As Boolean
            For Each pair In session.EditableEntriesByLine
                Dim lineNumber = pair.Key
                Dim entry = pair.Value
                Dim originalLine As String = Nothing
                If Not session.OriginalLinesByLine.TryGetValue(lineNumber, originalLine) Then
                    Continue For
                End If

                Dim editedLine = NormalizeInlineEditSingleLine(If(entry?.InlineEditText, String.Empty))
                If Not StringComparer.Ordinal.Equals(editedLine, originalLine) Then
                    hasAnyChange = True
                    Exit For
                End If
            Next

            If Not hasAnyChange Then
                ShowStatus("No inline edits to save.")
                UpdateGitDiffInlineToolbarState()
                Return
            End If

            _gitInlineEditSaveInProgress = True
            UpdateGitDiffInlineToolbarState()

            Try
                Dim fileBuffer = Await Task.Run(Function() ReadEditableTextFileBuffer(session.FullPath)).ConfigureAwait(True)
                If fileBuffer.Lines Is Nothing Then
                    fileBuffer.Lines = New List(Of String)()
                End If

                Dim lineNumbers As New List(Of Integer)(session.EditableEntriesByLine.Keys)
                lineNumbers.Sort()

                For Each lineNumber In lineNumbers
                    If lineNumber <= 0 OrElse lineNumber > fileBuffer.Lines.Count Then
                        Throw New InvalidOperationException("Selected line range no longer exists in the file.")
                    End If

                    Dim expectedLine As String = Nothing
                    If Not session.OriginalLinesByLine.TryGetValue(lineNumber, expectedLine) Then
                        Throw New InvalidOperationException("Inline edit state is out of date.")
                    End If

                    Dim currentLine = fileBuffer.Lines(lineNumber - 1)
                    If Not StringComparer.Ordinal.Equals(currentLine, expectedLine) Then
                        Throw New InvalidOperationException("The file changed since you opened inline edit. Refresh the diff and try again.")
                    End If

                    Dim entry = session.EditableEntriesByLine(lineNumber)
                    Dim editedLine = NormalizeInlineEditSingleLine(If(entry?.InlineEditText, String.Empty))
                    fileBuffer.Lines(lineNumber - 1) = editedLine
                Next

                fileBuffer.HadTerminalNewLine = session.HadTerminalNewLine
                fileBuffer.PreferredNewLine = If(String.IsNullOrEmpty(session.PreferredNewLine), vbLf, session.PreferredNewLine)
                fileBuffer.FileEncoding = If(session.FileEncoding, Encoding.UTF8)

                Await Task.Run(Sub() WriteEditableTextFileBuffer(session.FullPath, fileBuffer)).ConfigureAwait(True)

                CancelGitInlineDiffEdit(shouldShowStatus:=False, force:=True)
                ShowStatus($"Updated {session.DisplayPath} ({lineNumbers.Count} line(s)).", displayToast:=True)
                Await RefreshGitPanelAsync().ConfigureAwait(True)
            Catch ex As Exception
                ShowStatus($"Could not save inline diff edits: {ex.Message}", isError:=True, displayToast:=True)
            Finally
                _gitInlineEditSaveInProgress = False
                UpdateGitDiffInlineToolbarState()
            End Try
        End Function

        Private Sub CancelGitInlineDiffEdit(Optional shouldShowStatus As Boolean = True, Optional force As Boolean = False)
            If _gitInlineEditSaveInProgress AndAlso Not force Then
                Return
            End If

            EndGitDiffCtrlDragSelection()
            Dim activeSession = _gitInlineDiffEditSession
            _gitInlineDiffEditSession = Nothing

            If activeSession IsNot Nothing AndAlso activeSession.EditableEntriesByLine IsNot Nothing Then
                For Each pair In activeSession.EditableEntriesByLine
                    Dim lineEntry = pair.Value
                    If lineEntry Is Nothing Then
                        Continue For
                    End If

                    lineEntry.IsInlineEditing = False
                    lineEntry.InlineEditText = String.Empty
                Next
            End If

            If GitPaneHost Is Nothing Then
                Return
            End If

            RefreshGitDiffPreviewListBindings()
            UpdateGitDiffInlineToolbarState()

            If shouldShowStatus Then
                ShowStatus("Canceled inline diff edit.")
            End If
        End Sub

        Private Sub EndGitDiffCtrlDragSelection()
            If Not _gitDiffCtrlDragSelecting Then
                _gitDiffCtrlDragAnchorIndex = -1
                _gitDiffCtrlDragLastIndex = -1
                Return
            End If

            _gitDiffCtrlDragSelecting = False
            _gitDiffCtrlDragAnchorIndex = -1
            _gitDiffCtrlDragLastIndex = -1

            Dim list = GitPaneHost?.LstGitPanelDiffPreviewLines
            If list IsNot Nothing AndAlso Mouse.Captured Is list Then
                Mouse.Capture(Nothing)
            End If
        End Sub

        Private Sub EnsureGitDiffLineSelection(startIndex As Integer, endIndex As Integer)
            Dim list = GitPaneHost?.LstGitPanelDiffPreviewLines
            If list Is Nothing OrElse list.Items Is Nothing OrElse list.Items.Count = 0 Then
                Return
            End If

            Dim minIndex = Math.Max(0, Math.Min(startIndex, endIndex))
            Dim maxIndex = Math.Min(list.Items.Count - 1, Math.Max(startIndex, endIndex))
            For i = minIndex To maxIndex
                Dim item = list.Items(i)
                If item Is Nothing Then
                    Continue For
                End If

                If Not list.SelectedItems.Contains(item) Then
                    list.SelectedItems.Add(item)
                End If
            Next
        End Sub

        Private Function ResolveGitDiffListItemIndex(listItem As ListBoxItem) As Integer
            Dim list = GitPaneHost?.LstGitPanelDiffPreviewLines
            If list Is Nothing OrElse listItem Is Nothing Then
                Return -1
            End If

            Dim dataItem = list.ItemContainerGenerator.ItemFromContainer(listItem)
            If dataItem Is DependencyProperty.UnsetValue OrElse dataItem Is Nothing Then
                dataItem = listItem.DataContext
            End If

            If dataItem Is Nothing Then
                Return -1
            End If

            Return list.Items.IndexOf(dataItem)
        End Function

        Private Function ResolveGitDiffListItemIndexAtPoint(list As ListBox, point As Point) As Integer
            If list Is Nothing Then
                Return -1
            End If

            Dim hit = TryCast(list.InputHitTest(point), DependencyObject)
            If hit Is Nothing Then
                Return -1
            End If

            If FindVisualAncestor(Of ScrollBar)(hit) IsNot Nothing Then
                Return -1
            End If

            Dim listItem = FindVisualAncestor(Of ListBoxItem)(hit)
            If listItem Is Nothing Then
                Return -1
            End If

            Return ResolveGitDiffListItemIndex(listItem)
        End Function

        Private Sub RestoreGitPanelDiffMetaForCurrentSelection()
            If GitPaneHost Is Nothing OrElse GitPaneHost.LblGitPanelDiffMeta Is Nothing Then
                Return
            End If

            Dim selected = TryCast(GitPaneHost.LstGitPanelChanges?.SelectedItem, GitChangedFileListEntry)
            Dim metaParts As New List(Of String)()
            If selected IsNot Nothing Then
                If selected.AddedLineCount.HasValue AndAlso selected.AddedLineCount.Value > 0 Then
                    metaParts.Add($"+{selected.AddedLineCount.Value}")
                End If
                If selected.RemovedLineCount.HasValue AndAlso selected.RemovedLineCount.Value > 0 Then
                    metaParts.Add($"-{selected.RemovedLineCount.Value}")
                End If
            End If

            If _gitInlineDiffEditSession IsNot Nothing Then
                metaParts.Add($"editing {_gitInlineDiffEditSession.EditableEntriesByLine.Count} line(s)")
            Else
                Dim selectedEditableCount = ResolveSelectedEditableGitDiffEntries().Count
                If selectedEditableCount > 0 Then
                    metaParts.Add($"{selectedEditableCount} selected")
                End If
            End If

            GitPaneHost.LblGitPanelDiffMeta.Text = String.Join("  ", metaParts)
        End Sub

        Private Sub UpdateGitDiffInlineToolbarState()
            If GitPaneHost Is Nothing Then
                Return
            End If

            Dim isEditing = _gitInlineDiffEditSession IsNot Nothing
            Dim hasEditableSelection = ResolveSelectedEditableGitDiffEntries().Count > 0
            Dim canStartEditing = Not isEditing AndAlso
                                  Not _gitInlineEditSaveInProgress AndAlso
                                  Not _gitPanelLoadingActive AndAlso
                                  Not _gitPanelCommandInProgress AndAlso
                                  hasEditableSelection
            Dim canSaveOrCancel = isEditing AndAlso Not _gitInlineEditSaveInProgress

            If GitPaneHost.BtnGitDiffInlineStart IsNot Nothing Then
                GitPaneHost.BtnGitDiffInlineStart.Visibility = If(isEditing, Visibility.Collapsed, Visibility.Visible)
                GitPaneHost.BtnGitDiffInlineStart.IsEnabled = canStartEditing
            End If

            If GitPaneHost.BtnGitDiffInlineSave IsNot Nothing Then
                GitPaneHost.BtnGitDiffInlineSave.Visibility = If(isEditing, Visibility.Visible, Visibility.Collapsed)
                GitPaneHost.BtnGitDiffInlineSave.IsEnabled = canSaveOrCancel
            End If

            If GitPaneHost.BtnGitDiffInlineCancel IsNot Nothing Then
                GitPaneHost.BtnGitDiffInlineCancel.Visibility = If(isEditing, Visibility.Visible, Visibility.Collapsed)
                GitPaneHost.BtnGitDiffInlineCancel.IsEnabled = canSaveOrCancel
            End If

            RestoreGitPanelDiffMetaForCurrentSelection()
        End Sub

        Private Sub RefreshGitDiffPreviewListBindings()
            If GitPaneHost Is Nothing OrElse GitPaneHost.LstGitPanelDiffPreviewLines Is Nothing Then
                Return
            End If

            Dim view = CollectionViewSource.GetDefaultView(GitPaneHost.LstGitPanelDiffPreviewLines.ItemsSource)
            If view IsNot Nothing Then
                view.Refresh()
            End If
        End Sub

        Private Sub FocusFirstGitInlineEditor()
            Dim list = GitPaneHost?.LstGitPanelDiffPreviewLines
            If list Is Nothing OrElse _gitInlineDiffEditSession Is Nothing Then
                Return
            End If

            Dispatcher.BeginInvoke(
                New Action(
                    Sub()
                        If _gitInlineDiffEditSession Is Nothing Then
                            Return
                        End If

                        For Each lineEntry In _gitInlineDiffEditSession.EditableEntriesByLine.Values
                            If lineEntry Is Nothing Then
                                Continue For
                            End If

                            list.ScrollIntoView(lineEntry)
                            Dim container = TryCast(list.ItemContainerGenerator.ContainerFromItem(lineEntry), ListBoxItem)
                            If container Is Nothing Then
                                Continue For
                            End If

                            Dim editor = FindVisualDescendantByName(Of TextBox)(container, "GitDiffInlineEditTextBox")
                            If editor Is Nothing Then
                                Continue For
                            End If

                            editor.Focus()
                            editor.SelectAll()
                            Exit For
                        Next
                    End Sub),
                DispatcherPriority.Input)
        End Sub

        Private Shared Function NormalizeInlineEditSingleLine(value As String) As String
            Dim normalized = If(value, String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim newlineIndex = normalized.IndexOf(vbLf, StringComparison.Ordinal)
            If newlineIndex >= 0 Then
                normalized = normalized.Substring(0, newlineIndex)
            End If

            Return normalized
        End Function

        Private Shared Function ReadEditableTextFileBuffer(fullPath As String) As EditableTextFileBuffer
            Dim text As String = String.Empty
            Dim encoding As Encoding = New UTF8Encoding(False)
            Using reader As New StreamReader(fullPath, detectEncodingFromByteOrderMarks:=True)
                text = reader.ReadToEnd()
                encoding = reader.CurrentEncoding
            End Using

            Dim preferredNewLine = If(text.Contains(vbCrLf), vbCrLf, vbLf)
            Dim normalized = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim hadTerminalNewLine = normalized.EndsWith(vbLf, StringComparison.Ordinal)
            Dim lines As New List(Of String)()

            If normalized.Length > 0 Then
                lines.AddRange(normalized.Split({vbLf}, StringSplitOptions.None))
                If hadTerminalNewLine AndAlso lines.Count > 0 Then
                    lines.RemoveAt(lines.Count - 1)
                End If
            End If

            Return New EditableTextFileBuffer With {
                .Lines = lines,
                .FileEncoding = encoding,
                .PreferredNewLine = preferredNewLine,
                .HadTerminalNewLine = hadTerminalNewLine
            }
        End Function

        Private Shared Sub WriteEditableTextFileBuffer(fullPath As String, buffer As EditableTextFileBuffer)
            Dim lines = If(buffer.Lines, New List(Of String)())
            Dim normalized = String.Join(vbLf, lines)
            If buffer.HadTerminalNewLine Then
                normalized &= vbLf
            End If

            Dim preferredNewLine = If(String.IsNullOrEmpty(buffer.PreferredNewLine), vbLf, buffer.PreferredNewLine)
            Dim content = If(StringComparer.Ordinal.Equals(preferredNewLine, vbCrLf),
                             normalized.Replace(vbLf, vbCrLf),
                             normalized)

            File.WriteAllText(fullPath, content, If(buffer.FileEncoding, New UTF8Encoding(False)))
        End Sub

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
                CloseGitPanel(playToggleSound:=True)
                Return
            End If

            HideMetricsPanelForPanelSwitch()
            ShowGitPanelDock()
            GitPaneHost.GitInspectorPanel.Visibility = Visibility.Visible
            UpdateSidebarSelectionState(showSettings:=(_viewModel.SidebarSettingsViewVisibility = Visibility.Visible))
            PlayPanelVisibilityToggleSoundIfEnabled()
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

        Private Sub CloseGitPanel(Optional playToggleSound As Boolean = True)
            If GitPaneHost.GitInspectorPanel Is Nothing Then
                Return
            End If

            If GitPaneHost.GitInspectorPanel.Visibility <> Visibility.Visible Then
                Return
            End If

            EndGitDiffCtrlDragSelection()
            CancelGitInlineDiffEdit(shouldShowStatus:=False, force:=True)

            Dim keepDockOpenForMetrics = IsMetricsPanelVisible()
            Dim actualWidth = GitPaneHost.GitInspectorPanel.ActualWidth
            If Not Double.IsNaN(actualWidth) AndAlso Not Double.IsInfinity(actualWidth) AndAlso actualWidth >= 280 Then
                _gitPanelDockWidth = actualWidth
            End If

            GitPaneHost.GitInspectorPanel.Visibility = Visibility.Collapsed
            If Not keepDockOpenForMetrics Then
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
            End If
            UpdateMainPaneResizeBounds()
            Interlocked.Increment(_gitPanelLoadVersion)
            Interlocked.Increment(_gitPanelDiffPreviewLoadVersion)
            Interlocked.Increment(_gitPanelCommitPreviewLoadVersion)
            Interlocked.Increment(_gitPanelBranchPreviewLoadVersion)
            UpdateSidebarSelectionState(showSettings:=(_viewModel.SidebarSettingsViewVisibility = Visibility.Visible))
            If playToggleSound Then
                PlayPanelVisibilityToggleSoundIfEnabled()
            End If
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
            _gitPanelLoadingActive = isLoading

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
                GitPaneHost.BtnGitPanelRefresh.IsEnabled = Not isLoading AndAlso Not _gitPanelCommandInProgress
            End If

            UpdateGitCommitComposerState()
            UpdateGitDiffInlineToolbarState()
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
                snapshot.StagedChangeCount = parsedStatus.StagedChangeCount
                snapshot.UnstagedChangeCount = parsedStatus.UnstagedChangeCount
                snapshot.UntrackedChangeCount = parsedStatus.UntrackedChangeCount
                snapshot.ConflictChangeCount = parsedStatus.ConflictChangeCount
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
            Public Property StagedChangeCount As Integer
            Public Property UnstagedChangeCount As Integer
            Public Property UntrackedChangeCount As Integer
            Public Property ConflictChangeCount As Integer
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

            result.StatusSummary = If(parts.Count = 0, "clean", String.Join(", ", parts))
            result.StagedChangeCount = stagedCount
            result.UnstagedChangeCount = unstagedCount
            result.UntrackedChangeCount = untrackedCount
            result.ConflictChangeCount = conflictCount
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

            Return New GitChangedFileListEntry() With {
                .StatusCode = If(rawCode, String.Empty),
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

        Private Shared Function ResolveGitCommandSummary(result As ProcessCaptureResult, fallback As String) As String
            Dim summary = FirstNonEmptyLine(result.OutputText)
            If String.IsNullOrWhiteSpace(summary) Then
                summary = FirstNonEmptyLine(result.ErrorText)
            End If

            If String.IsNullOrWhiteSpace(summary) Then
                Return If(fallback, String.Empty)
            End If

            Return summary
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
            StopAndDisposeUiSoundPlayers()
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
                If TryExecuteTurnComposerSendOrSteerCommand() Then
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
            Dim normalizedStatusMessage = If(message, String.Empty).Trim()
            _viewModel.StatusText = normalizedStatusMessage
            StatusBarPaneHost.LblStatus.Foreground = If(isError,
                                      ResolveBrush("DangerBrush", Brushes.DarkRed),
                                      ResolveBrush("TextPrimaryBrush", Brushes.Black))

            If IsApprovalNeededStatusMessage(normalizedStatusMessage) Then
                PlayApprovalNeededSoundIfEnabled(normalizedStatusMessage)
            ElseIf ShouldPlayGeneralErrorStatusSound(normalizedStatusMessage, isError) Then
                PlayGeneralErrorSoundIfEnabled(normalizedStatusMessage)
            End If

            Dim suppressHintToast = displayToast AndAlso
                                    Not isError AndAlso
                                    _viewModel IsNot Nothing AndAlso
                                    _viewModel.SettingsPanel IsNot Nothing AndAlso
                                    _viewModel.SettingsPanel.DisableConnectionInitializedToast

            If displayToast AndAlso Not suppressHintToast Then
                ShowToast(normalizedStatusMessage, isError)
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

            Dim hasActiveTurn = DetermineHasActiveTurn(session)
            _viewModel.TranscriptPanel.SetActiveTurnRetentionEnabled(hasActiveTurn)
            UpdateActiveTurnProgressIndicatorUi(hasActiveTurn, advanceWorkingDots:=False)
            Dim authenticatedAndInteractive = authenticated AndAlso
                                             Not _logoutUiTransitionInProgress AndAlso
                                             Not _disconnectUiTransitionInProgress
            Dim canUseExistingThreadTurnControls = authenticatedAndInteractive AndAlso Not _threadContentLoading AndAlso session.HasCurrentThread
            Dim canStartTurn = authenticatedAndInteractive AndAlso Not _threadContentLoading

            _viewModel.TurnComposer.CanStartTurn = canStartTurn AndAlso Not hasActiveTurn
            _viewModel.TurnComposer.CanSteerTurn = canUseExistingThreadTurnControls AndAlso hasActiveTurn
            _viewModel.TurnComposer.CanInterruptTurn = canUseExistingThreadTurnControls AndAlso hasActiveTurn
            _viewModel.TurnComposer.StartTurnVisibility = If(_viewModel.TurnComposer.CanInterruptTurn, Visibility.Collapsed, Visibility.Visible)
            _viewModel.TurnComposer.InterruptTurnVisibility = If(_viewModel.TurnComposer.CanInterruptTurn, Visibility.Visible, Visibility.Collapsed)
            _viewModel.TurnComposer.IsInputEnabled = authenticatedAndInteractive
            _viewModel.TurnComposer.IsModelEnabled = authenticatedAndInteractive
            _viewModel.TurnComposer.IsReasoningEnabled = authenticatedAndInteractive
            _viewModel.TurnComposer.IsApprovalPolicyEnabled = authenticatedAndInteractive
            _viewModel.TurnComposer.IsSandboxEnabled = authenticatedAndInteractive
            If Not authenticatedAndInteractive Then
                HideTurnComposerTokenSuggestionsPopup()
            ElseIf WorkspacePaneHost IsNot Nothing AndAlso WorkspacePaneHost.TxtTurnInput IsNot Nothing AndAlso
                   WorkspacePaneHost.TxtTurnInput.IsKeyboardFocusWithin Then
                RefreshTurnComposerTokenSuggestionsPopup(triggerCatalogWarmup:=False)
            End If

            _viewModel.IsSidebarNewThreadEnabled = authenticatedAndInteractive AndAlso Not _threadContentLoading
            _viewModel.IsSidebarMetricsEnabled = True
            _viewModel.IsSidebarAutomationsEnabled = True
            _viewModel.IsSidebarSkillsEnabled = True
            _viewModel.IsSidebarSettingsEnabled = True
            _viewModel.IsSettingsBackEnabled = True
            _viewModel.IsQuickOpenVscEnabled = True
            _viewModel.IsQuickOpenTerminalEnabled = True
            SetTranscriptTabInteractionEnabled(authenticatedAndInteractive)
            Dim hasLiveOverlayHistoryForCurrentThread = False
            Dim visibleThreadId = GetVisibleThreadId()
            If Not String.IsNullOrWhiteSpace(visibleThreadId) Then
                hasLiveOverlayHistoryForCurrentThread = _threadLiveSessionRegistry.GetOverlayTurnIds(visibleThreadId).Count > 0
            End If

            _viewModel.TranscriptPanel.CollapseCommandDetailsByDefault = hasActiveTurn OrElse hasLiveOverlayHistoryForCurrentThread
        End Sub

        Private Function DetermineHasActiveTurn(session As SessionStateViewModel) As Boolean
            Dim hasActiveTurn = HasActiveRuntimeTurnForCurrentThread()
            If Not hasActiveTurn AndAlso session IsNot Nothing AndAlso session.HasCurrentTurn AndAlso
               Not RuntimeHasTurnHistoryForCurrentThread() Then
                hasActiveTurn = True
            End If

            Return hasActiveTurn
        End Function

        Private Sub UpdateActiveTurnProgressIndicatorUi(hasActiveTurn As Boolean,
                                                        Optional advanceWorkingDots As Boolean = False)
            If _viewModel Is Nothing OrElse _viewModel.TranscriptPanel Is Nothing Then
                Return
            End If

            Dim transcriptPanel = _viewModel.TranscriptPanel
            If Not hasActiveTurn Then
                transcriptPanel.ActiveTurnSpinnerVisibility = Visibility.Collapsed
                transcriptPanel.ActiveTurnElapsedText = "00:00"
                transcriptPanel.ActiveTurnStatusText = "Working..."
                _activeTurnProgressTurnId = String.Empty
                _activeTurnProgressStartedAtUtc = Nothing
                _activeTurnProgressWorkingDotsPhase = 0
                Return
            End If

            transcriptPanel.ActiveTurnSpinnerVisibility = Visibility.Visible

            Dim threadId = GetVisibleThreadId()
            Dim turnId = GetVisibleTurnId()
            Dim activeTurnId = GetActiveTurnIdForThread(threadId, turnId)
            If String.IsNullOrWhiteSpace(activeTurnId) Then
                activeTurnId = turnId
            End If

            If Not StringComparer.Ordinal.Equals(_activeTurnProgressTurnId, activeTurnId) Then
                _activeTurnProgressTurnId = If(activeTurnId, String.Empty).Trim()
                _activeTurnProgressStartedAtUtc = Nothing
                _activeTurnProgressWorkingDotsPhase = 3
            End If

            Dim resolvedStartUtc = ResolveActiveTurnStartedAtUtc(threadId, activeTurnId)
            If resolvedStartUtc.HasValue Then
                _activeTurnProgressStartedAtUtc = resolvedStartUtc
            ElseIf Not _activeTurnProgressStartedAtUtc.HasValue Then
                _activeTurnProgressStartedAtUtc = DateTimeOffset.UtcNow
            End If

            Dim elapsed = DateTimeOffset.UtcNow - _activeTurnProgressStartedAtUtc.Value
            transcriptPanel.ActiveTurnElapsedText = FormatActiveTurnElapsed(elapsed)

            Dim reasoningText = transcriptPanel.GetLatestReasoningCardStatusText(activeTurnId)
            transcriptPanel.ActiveTurnStatusText = BuildActiveTurnStatusText(reasoningText, advanceWorkingDots)
        End Sub

        Private Function ResolveActiveTurnStartedAtUtc(threadId As String, turnId As String) As DateTimeOffset?
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            Dim normalizedTurnId = If(turnId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return Nothing
            End If

            Dim runtimeStore = _sessionNotificationCoordinator?.RuntimeStore
            If runtimeStore Is Nothing Then
                Return Nothing
            End If

            Dim turnState = runtimeStore.GetTurnState(normalizedThreadId, normalizedTurnId)
            If turnState Is Nothing OrElse Not turnState.StartedAt.HasValue Then
                Return Nothing
            End If

            Return turnState.StartedAt.Value
        End Function

        Private Shared Function FormatActiveTurnElapsed(elapsed As TimeSpan) As String
            Dim totalSeconds = Math.Max(0, CInt(Math.Floor(elapsed.TotalSeconds)))
            Dim minutes = totalSeconds \ 60
            Dim seconds = totalSeconds Mod 60
            Return $"{minutes:00}:{seconds:00}"
        End Function

        Private Function BuildActiveTurnStatusText(reasoningText As String, advanceWorkingDots As Boolean) As String
            Dim normalizedReasoning = If(reasoningText, String.Empty).Trim()

            If _activeTurnProgressWorkingDotsPhase < 1 OrElse _activeTurnProgressWorkingDotsPhase > 3 Then
                _activeTurnProgressWorkingDotsPhase = 3
            End If

            If advanceWorkingDots Then
                _activeTurnProgressWorkingDotsPhase = (_activeTurnProgressWorkingDotsPhase Mod 3) + 1
            End If

            Dim animatedDots = New String("."c, _activeTurnProgressWorkingDotsPhase)
            If Not String.IsNullOrWhiteSpace(normalizedReasoning) Then
                Return normalizedReasoning & animatedDots
            End If

            Return "Working" & animatedDots
        End Function

        Private Sub RefreshApprovalControlState(authenticated As Boolean)
            Dim visibleThreadId = GetVisibleThreadId()
            If _turnWorkflowCoordinator Is Nothing Then
                _viewModel.ApprovalPanel.SetThreadScopedState(String.Empty,
                                                              String.Empty,
                                                              supportsExecpolicyAmendment:=False,
                                                              pendingQueueCount:=0)
                _viewModel.ApprovalPanel.UpdateAvailability(authenticated, hasActiveApproval:=False)
                Return
            End If

            _turnWorkflowCoordinator.RefreshApprovalPanelForThread(visibleThreadId, authenticated)
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

            If String.IsNullOrWhiteSpace(GetVisibleThreadId()) Then
                Return "Select a thread from the left panel, or send your first instruction to start a new one."
            End If

            If String.IsNullOrWhiteSpace(GetVisibleTurnId()) Then
                Return "Ready. Press Enter to send."
            End If

            Return "Turn in progress. Press Enter to steer or use Interrupt to stop execution."
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
            Dim showMetrics = Not showSettings AndAlso IsMetricsPanelVisible()
            Dim newThreadTag As String
            If showSettings OrElse showMetrics Then
                newThreadTag = String.Empty
            Else
                newThreadTag = "Active"
            End If

            _viewModel.SidebarNewThreadNavTag = newThreadTag
            _viewModel.SidebarMetricsNavTag = If(showMetrics, "Active", String.Empty)
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
            settingsObject("showTurnLifecycleDotsInTranscript") = _viewModel.SettingsPanel.ShowTurnLifecycleDotsInTranscript
            settingsObject("showReasoningBubblesInTranscript") = _viewModel.SettingsPanel.ShowReasoningBubblesInTranscript
            settingsObject("playUiSounds") = _viewModel.SettingsPanel.PlayUiSounds
            settingsObject("uiSoundVolumePercent") = _viewModel.SettingsPanel.UiSoundVolumePercent
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

            If Not IsTranscriptListRootScrollViewer(scroller) Then
                DebugTranscriptScroll("scroll_changed", "skip=nested_scrollviewer", scroller, e)
                Return
            End If

            _transcriptScrollViewer = scroller

            If _transcriptScrollProgrammaticMoveInProgress Then
                DebugTranscriptScroll("scroll_changed", "skip=programmatic_move", scroller, e)
                Return
            End If

            Dim atBottom = IsScrollViewerAtBottomExtreme(scroller)
            Dim layoutMetricsChanged = Math.Abs(e.ExtentHeightChange) > 0.01R OrElse
                                      Math.Abs(e.ViewportHeightChange) > 0.01R

            Dim verticalOffsetChanged = Math.Abs(e.VerticalChange) > 0.01R
            If Not verticalOffsetChanged Then
                ' Consume one-shot user-arm on no-op clicks/wheel/key events so later content-growth
                ' ScrollChanged events are not misclassified as user navigation.
                If _transcriptUserScrollInteractionArmed AndAlso
                   Not _transcriptScrollThumbDragActive AndAlso
                   Not layoutMetricsChanged Then
                    _transcriptUserScrollInteractionArmed = False
                End If

                If atBottom AndAlso _transcriptScrollThumbDragActive Then
                    SetTranscriptFollowModeFollowBottom()
                End If
                DebugTranscriptScroll("scroll_changed", $"skip=no_vertical_change;layoutMetricsChanged={layoutMetricsChanged}", scroller, e)
                Return
            End If

            If _transcriptUserScrollInteractionArmed AndAlso
               Not _transcriptScrollThumbDragActive AndAlso
               IsTranscriptFollowBottomEnabled() AndAlso
               layoutMetricsChanged AndAlso
               e.ExtentHeightChange > 0.01R AndAlso
               e.VerticalChange >= -0.01R Then
                ' Content growth while pinned at bottom can move the offset and extent in the same event.
                ' If a no-op user input (e.g. wheel-down at bottom) armed interaction, do not misclassify
                ' this as a user navigation detach.
                _transcriptUserScrollInteractionArmed = False
                DebugTranscriptScroll("scroll_changed", "skip=armed_content_growth_while_following", scroller, e)
                Return
            End If

            Dim isUserDrivenOffsetChange = _transcriptScrollThumbDragActive OrElse
                                           _transcriptUserScrollInteractionArmed
            If Not isUserDrivenOffsetChange Then
                If layoutMetricsChanged AndAlso
                   _transcriptScrollFollowMode = TranscriptScrollFollowMode.DetachedByUser AndAlso
                   TryRestoreTranscriptDetachedAnchorOffset(scroller) Then
                    DebugTranscriptScroll("scroll_changed", "restore=detached_anchor_after_layout_change", scroller, e)
                    Return
                End If

                DebugTranscriptScroll("scroll_changed", "skip=content_or_layout_change", scroller, e)
                Return
            End If

            If Not atBottom Then
                CancelPendingTranscriptScrollRequests()
            End If

            If _transcriptScrollThumbDragActive Then
                SetTranscriptFollowModeDetachedByUser()
            ElseIf atBottom AndAlso Not layoutMetricsChanged AndAlso IsScrollViewerAtStableBottomForUserReattach(scroller) Then
                SetTranscriptFollowModeFollowBottom()
            Else
                SetTranscriptFollowModeDetachedByUser()
            End If

            If _transcriptScrollFollowMode = TranscriptScrollFollowMode.DetachedByUser Then
                UpdateTranscriptDetachedAnchorOffset(scroller)
            End If

            If Not _transcriptScrollThumbDragActive Then
                _transcriptUserScrollInteractionArmed = False
            End If

            DebugTranscriptScroll("scroll_changed", "applied=user_navigation_state_update", scroller, e)
        End Sub

        Private Sub OnTranscriptPreviewMouseWheel(sender As Object, e As MouseWheelEventArgs)
            If e Is Nothing OrElse _suppressTranscriptScrollTracking Then
                Return
            End If

            Dim scroller = ResolveTranscriptScrollViewer()
            Dim atBottom = IsScrollViewerAtBottomExtreme(scroller)
            Dim detachFollow = e.Delta > 0 OrElse Not atBottom

            If e.Delta < 0 AndAlso atBottom AndAlso IsTranscriptFollowBottomEnabled() Then
                DebugTranscriptScroll("mouse_wheel", $"skip=noop_pinned_bottom;delta={e.Delta}", scroller)
                Return
            End If

            DebugTranscriptScroll("mouse_wheel", $"delta={e.Delta};detachFollow={detachFollow}", scroller)
            BeginTranscriptUserScrollInteraction(detachFollow)
        End Sub

        Private Sub OnTranscriptPreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
            If e Is Nothing OrElse _suppressTranscriptScrollTracking Then
                Return
            End If

            Dim source = TryCast(e.OriginalSource, DependencyObject)
            If source Is Nothing Then
                DebugTranscriptScroll("mouse_left_down", "skip=no_source")
                Return
            End If

            Dim scrollBar = FindVisualAncestor(Of ScrollBar)(source)
            If scrollBar Is Nothing Then
                DebugTranscriptScroll("mouse_left_down", "skip=not_scrollbar_click")
                Return
            End If

            If Not IsTranscriptScrollbar(scrollBar) Then
                DebugTranscriptScroll("mouse_left_down", "skip=non_transcript_scrollbar")
                Return
            End If

            DebugTranscriptScroll("mouse_left_down", "scrollbar_click")
            BeginTranscriptUserScrollInteraction(detachFollow:=True)
        End Sub

        Private Sub OnTranscriptPreviewKeyDown(sender As Object, e As KeyEventArgs)
            If e Is Nothing OrElse _suppressTranscriptScrollTracking Then
                Return
            End If

            Select Case e.Key
                Case Key.Up, Key.PageUp, Key.Home, Key.Down, Key.PageDown, Key.End
                    Dim scroller = ResolveTranscriptScrollViewer()
                    Dim atBottom = IsScrollViewerAtBottomExtreme(scroller)
                    If (e.Key = Key.Down OrElse e.Key = Key.PageDown OrElse e.Key = Key.End) AndAlso
                       atBottom AndAlso
                       IsTranscriptFollowBottomEnabled() Then
                        DebugTranscriptScroll("key_down", $"skip=noop_pinned_bottom;key={e.Key}", scroller)
                        Return
                    End If

                    Dim detachFollow = (e.Key = Key.Up OrElse e.Key = Key.PageUp OrElse e.Key = Key.Home)
                    DebugTranscriptScroll("key_down", $"key={e.Key};detachFollow={detachFollow}")
                    BeginTranscriptUserScrollInteraction(detachFollow)
            End Select
        End Sub

        Private Sub OnTranscriptScrollThumbDragStarted(sender As Object, e As DragStartedEventArgs)
            If e Is Nothing OrElse _suppressTranscriptScrollTracking Then
                Return
            End If

            Dim thumb = TryCast(e.OriginalSource, DependencyObject)
            If thumb Is Nothing Then
                DebugTranscriptScroll("thumb_drag_started", "skip=no_source")
                Return
            End If

            Dim scrollBar = FindVisualAncestor(Of ScrollBar)(thumb)
            If scrollBar Is Nothing OrElse Not IsTranscriptScrollbar(scrollBar) Then
                DebugTranscriptScroll("thumb_drag_started", "skip=non_transcript_scrollbar")
                Return
            End If

            _transcriptScrollThumbDragActive = True
            DebugTranscriptScroll("thumb_drag_started")
            BeginTranscriptUserScrollInteraction(detachFollow:=True)
        End Sub

        Private Sub OnTranscriptScrollThumbDragCompleted(sender As Object, e As DragCompletedEventArgs)
            If e Is Nothing OrElse _suppressTranscriptScrollTracking Then
                Return
            End If

            Dim thumb = TryCast(e.OriginalSource, DependencyObject)
            If thumb IsNot Nothing Then
                Dim scrollBar = FindVisualAncestor(Of ScrollBar)(thumb)
                If scrollBar Is Nothing OrElse Not IsTranscriptScrollbar(scrollBar) Then
                    DebugTranscriptScroll("thumb_drag_completed", "skip=non_transcript_scrollbar")
                    Return
                End If
            End If

            _transcriptScrollThumbDragActive = False

            Dim scroller = ResolveTranscriptScrollViewer()
            If scroller IsNot Nothing Then
                If IsScrollViewerAtBottomExtreme(scroller) Then
                    SetTranscriptFollowModeFollowBottom()
                Else
                    SetTranscriptFollowModeDetachedByUser()
                End If
            End If

            _transcriptUserScrollInteractionArmed = False
            DebugTranscriptScroll("thumb_drag_completed", "drag_end_state_applied", scroller)
        End Sub

        Private Sub BeginTranscriptUserScrollInteraction(Optional detachFollow As Boolean = True)
            _transcriptUserScrollInteractionArmed = True
            _transcriptScrollUserInteractionEpoch += 1
            CancelPendingTranscriptScrollRequests()
            If detachFollow Then
                SetTranscriptFollowModeDetachedByUser()
                UpdateTranscriptDetachedAnchorOffset(ResolveTranscriptScrollViewer())
            End If
            DebugTranscriptScroll("user_interaction_begin", $"detachFollow={detachFollow}")
        End Sub

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
            Dim noThreadSelected = String.IsNullOrWhiteSpace(GetVisibleThreadId())
            Dim noActiveTurn = String.IsNullOrWhiteSpace(GetVisibleTurnId())
            Dim showDraftNewThreadEmptyState = _pendingNewThreadFirstPromptSelection
            Dim showEmptyState = (noThreadSelected OrElse showDraftNewThreadEmptyState) AndAlso
                                 noActiveTurn AndAlso
                                 Not hasVisibleTranscriptEntries AndAlso
                                 Not transcriptLoading

            WorkspacePaneHost.ImgTimeToBuildEmptyState.Visibility = If(showEmptyState, Visibility.Visible, Visibility.Collapsed)
        End Sub

        Private Shared Function IsScrollViewerAtBottomExtreme(scroller As ScrollViewer) As Boolean
            If scroller Is Nothing Then
                Return True
            End If

            If scroller.ScrollableHeight <= 0 Then
                Return True
            End If

            ' Use the ScrollViewer's actual bottom extreme rather than a "near bottom" pixel heuristic.
            Return scroller.VerticalOffset >= scroller.ScrollableHeight
        End Function

        Private Shared Function IsScrollViewerAtStableBottomForUserReattach(scroller As ScrollViewer) As Boolean
            If scroller Is Nothing Then
                Return True
            End If

            If scroller.ScrollableHeight <= 0 Then
                Return True
            End If

            ' WPF can briefly report offset > scrollable during layout churn while streaming.
            ' Require a tighter equality check before treating this as an intentional user reattach.
            Return Math.Abs(scroller.VerticalOffset - scroller.ScrollableHeight) <= 0.5R
        End Function

        Private Sub ScrollTranscriptToBottom(Optional force As Boolean = False,
                                            Optional reason As TranscriptScrollRequestReason = TranscriptScrollRequestReason.LegacyCallsite)
            RequestTranscriptScroll(reason, force)
        End Sub

        Private Sub RequestTranscriptScroll(reason As TranscriptScrollRequestReason,
                                            Optional force As Boolean = False)
            If WorkspacePaneHost Is Nothing Then
                Return
            End If

            Dim policy = ResolveTranscriptScrollRequestPolicy(reason, force)
            If policy = TranscriptScrollRequestPolicy.None Then
                DebugTranscriptScroll("request_scroll", $"skip=none_policy;reason={reason};force={force}")
                Return
            End If

                If policy = TranscriptScrollRequestPolicy.FollowIfPinned Then
                    If _transcriptScrollThumbDragActive OrElse _transcriptUserScrollInteractionArmed Then
                        DebugTranscriptScroll("request_scroll", $"skip=user_interacting;reason={reason};policy={policy};force={force}")
                        Return
                    End If

                    If reason = TranscriptScrollRequestReason.RuntimeStream AndAlso
                       Not IsTranscriptFollowBottomEnabled() Then
                        Dim scroller = ResolveTranscriptScrollViewer()
                        If scroller IsNot Nothing AndAlso
                           IsScrollViewerAtBottomExtreme(scroller) AndAlso
                           IsScrollViewerAtStableBottomForUserReattach(scroller) Then
                            SetTranscriptFollowModeFollowBottom()
                            DebugTranscriptScroll("request_scroll", $"auto_reattach=bottom;reason={reason};policy={policy}", scroller)
                        Else
                            DebugTranscriptScroll("request_scroll", $"skip=detached_follow;reason={reason};policy={policy}", scroller)
                            Return
                        End If
                    End If
                End If

            _transcriptScrollQueuedPolicy = MergeTranscriptScrollRequestPolicy(_transcriptScrollQueuedPolicy, policy)
            _transcriptScrollQueuedReasons = _transcriptScrollQueuedReasons Or reason
            _transcriptScrollQueuedInteractionEpoch = _transcriptScrollUserInteractionEpoch
            _transcriptScrollRequestGeneration += 1

            DebugTranscriptScroll("request_scroll", $"queued;reason={reason};force={force};policy={policy}")

            QueueTranscriptScrollRequestProcessing()
        End Sub

        Private Sub QueueTranscriptScrollRequestProcessing()
            If _transcriptScrollRequestPending Then
                DebugTranscriptScroll("queue_request_processing", "skip=already_pending")
                Return
            End If

            _transcriptScrollRequestPending = True
            DebugTranscriptScroll("queue_request_processing", "scheduled")
            Dispatcher.BeginInvoke(
                DispatcherPriority.Background,
                New Action(AddressOf ProcessTranscriptScrollRequestQueue))
        End Sub

        Private Sub ProcessTranscriptScrollRequestQueue()
            _transcriptScrollRequestPending = False

            Dim queuedPolicy = _transcriptScrollQueuedPolicy
            Dim queuedReasons = _transcriptScrollQueuedReasons
            Dim queuedInteractionEpoch = _transcriptScrollQueuedInteractionEpoch
            Dim requestGeneration = _transcriptScrollRequestGeneration

            _transcriptScrollQueuedPolicy = TranscriptScrollRequestPolicy.None
            _transcriptScrollQueuedReasons = TranscriptScrollRequestReason.None
            _transcriptScrollQueuedInteractionEpoch = 0

            If queuedPolicy = TranscriptScrollRequestPolicy.None Then
                DebugTranscriptScroll("process_request_queue", "skip=none_policy")
                Return
            End If

            DebugTranscriptScroll("process_request_queue", $"dispatch;policy={queuedPolicy};reasons={queuedReasons};epoch={queuedInteractionEpoch}")
            ApplyTranscriptScrollRequest(queuedPolicy, queuedReasons, queuedInteractionEpoch, requestGeneration)
        End Sub

        Private Sub ApplyTranscriptScrollRequest(policy As TranscriptScrollRequestPolicy,
                                                 reasons As TranscriptScrollRequestReason,
                                                 requestInteractionEpoch As Integer,
                                                 requestGeneration As Integer)
            If WorkspacePaneHost Is Nothing Then
                Return
            End If

            UpdateWorkspaceEmptyStateVisibility()

            Dim transcriptList = CurrentTranscriptListControl()
            If transcriptList IsNot Nothing Then
                Dim scroller = ResolveTranscriptScrollViewer()
                Dim runtimeFollowOnly = IsRuntimeOnlyTranscriptScrollRequest(reasons)
                Dim isForceJump = (policy = TranscriptScrollRequestPolicy.ForceJump)
                DebugTranscriptScroll("apply_request", $"enter;policy={policy};reasons={reasons};runtimeFollowOnly={runtimeFollowOnly};requestEpoch={requestInteractionEpoch};requestGen={requestGeneration}", scroller)
                If isForceJump Then
                    SetTranscriptFollowModeFollowBottom()
                ElseIf policy = TranscriptScrollRequestPolicy.FollowIfPinned AndAlso Not IsTranscriptFollowBottomEnabled() Then
                    If Not runtimeFollowOnly Then
                        ScrollTextBoxToBottom(WorkspacePaneHost.TxtTranscript)
                    End If
                    DebugTranscriptScroll("apply_request", "skip=follow_detached", scroller)
                    Return
                End If

                If _transcriptScrollThumbDragActive AndAlso Not isForceJump Then
                    If Not runtimeFollowOnly Then
                        ScrollTextBoxToBottom(WorkspacePaneHost.TxtTranscript)
                    End If
                    DebugTranscriptScroll("apply_request", "skip=thumb_drag_active", scroller)
                    Return
                End If

                If policy = TranscriptScrollRequestPolicy.FollowIfPinned AndAlso _transcriptUserScrollInteractionArmed Then
                    If Not runtimeFollowOnly Then
                        ScrollTextBoxToBottom(WorkspacePaneHost.TxtTranscript)
                    End If
                    DebugTranscriptScroll("apply_request", "skip=user_interaction_armed", scroller)
                    Return
                End If

                If scroller IsNot Nothing Then
                    DebugTranscriptScroll("apply_request", "queue_scroll_to_bottom", scroller)
                    QueueTranscriptScrollToBottom(policy, requestGeneration, requestInteractionEpoch)
                ElseIf transcriptList.Items IsNot Nothing AndAlso transcriptList.Items.Count > 0 AndAlso isForceJump Then
                    DebugTranscriptScroll("apply_request", "fallback=scroll_into_view_last_item")
                    transcriptList.ScrollIntoView(transcriptList.Items(transcriptList.Items.Count - 1))
                End If
            End If

            If Not IsRuntimeOnlyTranscriptScrollRequest(reasons) OrElse policy = TranscriptScrollRequestPolicy.ForceJump Then
                DebugTranscriptScroll("apply_request", "scroll_hidden_raw_transcript")
                ScrollTextBoxToBottom(WorkspacePaneHost.TxtTranscript)
            End If
        End Sub

        Private Function ResolveTranscriptScrollViewer() As ScrollViewer
            Dim transcriptList = CurrentTranscriptListControl()
            If transcriptList Is Nothing Then
                _transcriptScrollViewer = Nothing
                Return Nothing
            End If

            If _transcriptScrollViewer IsNot Nothing AndAlso Not IsTranscriptListRootScrollViewer(_transcriptScrollViewer) Then
                _transcriptScrollViewer = Nothing
            End If

            If _transcriptScrollViewer Is Nothing Then
                _transcriptScrollViewer = FindVisualDescendant(Of ScrollViewer)(transcriptList)
                If _transcriptScrollViewer IsNot Nothing AndAlso Not IsTranscriptListRootScrollViewer(_transcriptScrollViewer) Then
                    _transcriptScrollViewer = Nothing
                End If
            End If

            Return _transcriptScrollViewer
        End Function

        Private Function IsTranscriptScrollbar(scrollBar As ScrollBar) As Boolean
            If scrollBar Is Nothing Then
                Return False
            End If

            Dim scroller = FindVisualAncestor(Of ScrollViewer)(scrollBar)
            Return IsTranscriptListRootScrollViewer(scroller)
        End Function

        Private Function IsTranscriptListRootScrollViewer(scroller As ScrollViewer) As Boolean
            Dim transcriptList = CurrentTranscriptListControl()
            If scroller Is Nothing OrElse transcriptList Is Nothing Then
                Return False
            End If

            If FindVisualAncestor(Of TextBox)(scroller) IsNot Nothing Then
                Return False
            End If

            Dim owningList = FindVisualAncestor(Of ListBox)(scroller)
            Return ReferenceEquals(owningList, transcriptList)
        End Function

        Private Sub QueueTranscriptScrollToBottom(policy As TranscriptScrollRequestPolicy,
                                                  requestGeneration As Integer,
                                                  requestInteractionEpoch As Integer)
            _transcriptScrollToBottomQueuedGeneration = requestGeneration
            _transcriptScrollToBottomQueuedPolicy = MergeTranscriptScrollRequestPolicy(_transcriptScrollToBottomQueuedPolicy, policy)
            _transcriptScrollToBottomQueuedInteractionEpoch = requestInteractionEpoch

            DebugTranscriptScroll("queue_to_bottom", $"enqueue;policy={policy};requestGen={requestGeneration};requestEpoch={requestInteractionEpoch}")

            If _transcriptScrollToBottomPending Then
                DebugTranscriptScroll("queue_to_bottom", "skip=already_pending")
                Return
            End If

            _transcriptScrollToBottomPending = True
            DebugTranscriptScroll("queue_to_bottom", "scheduled")
            Dispatcher.BeginInvoke(
                DispatcherPriority.Background,
                New Action(
                    Sub()
                        Dim queuedGeneration = _transcriptScrollToBottomQueuedGeneration
                        Dim queuedPolicy = _transcriptScrollToBottomQueuedPolicy
                        Dim queuedInteractionEpoch = _transcriptScrollToBottomQueuedInteractionEpoch

                        _transcriptScrollToBottomPending = False
                        _transcriptScrollToBottomQueuedPolicy = TranscriptScrollRequestPolicy.None

                        Dim scroller = ResolveTranscriptScrollViewer()
                        If scroller Is Nothing Then
                            DebugTranscriptScroll("queue_to_bottom_cb", "skip=no_scroller")
                            Return
                        End If

                        If queuedGeneration <> _transcriptScrollRequestGeneration Then
                            DebugTranscriptScroll("queue_to_bottom_cb", $"skip=stale_generation;queuedGen={queuedGeneration}")
                            Return
                        End If

                        Dim isForceJump = (queuedPolicy = TranscriptScrollRequestPolicy.ForceJump)
                        If queuedPolicy = TranscriptScrollRequestPolicy.None Then
                            DebugTranscriptScroll("queue_to_bottom_cb", "skip=none_policy", scroller)
                            Return
                        End If

                        If _transcriptScrollThumbDragActive Then
                            DebugTranscriptScroll("queue_to_bottom_cb", "skip=thumb_drag_active", scroller)
                            Return
                        End If

                        If Not isForceJump Then
                            If _transcriptUserScrollInteractionArmed Then
                                DebugTranscriptScroll("queue_to_bottom_cb", "skip=user_interaction_armed", scroller)
                                Return
                            End If

                            If queuedInteractionEpoch <> _transcriptScrollUserInteractionEpoch Then
                                DebugTranscriptScroll("queue_to_bottom_cb", $"skip=stale_user_epoch;queuedEpoch={queuedInteractionEpoch}", scroller)
                                Return
                            End If

                            If Not IsTranscriptFollowBottomEnabled() Then
                                DebugTranscriptScroll("queue_to_bottom_cb", "skip=follow_detached", scroller)
                                Return
                            End If
                        End If

                        If IsScrollViewerAtBottomExtreme(scroller) Then
                            DebugTranscriptScroll("queue_to_bottom_cb", "skip=already_bottom", scroller)
                            Return
                        End If

                        DebugTranscriptScroll("queue_to_bottom_cb", $"scroll_to_bottom;policy={queuedPolicy};queuedGen={queuedGeneration};queuedEpoch={queuedInteractionEpoch}", scroller)
                        _transcriptScrollProgrammaticMoveInProgress = True
                        _suppressTranscriptScrollTracking = True
                        Try
                            scroller.ScrollToBottom()
                        Finally
                            _suppressTranscriptScrollTracking = False
                            _transcriptScrollProgrammaticMoveInProgress = False
                        End Try
                    End Sub))
        End Sub

        Private Function IsTranscriptFollowBottomEnabled() As Boolean
            Return _transcriptScrollFollowMode = TranscriptScrollFollowMode.FollowBottom
        End Function

        Private Sub SetTranscriptFollowModeFollowBottom()
            _transcriptScrollFollowMode = TranscriptScrollFollowMode.FollowBottom
            _transcriptDetachedAnchorOffset = Nothing
        End Sub

        Private Sub SetTranscriptFollowModeDetachedByUser()
            _transcriptScrollFollowMode = TranscriptScrollFollowMode.DetachedByUser
            If Not _transcriptDetachedAnchorOffset.HasValue Then
                UpdateTranscriptDetachedAnchorOffset(ResolveTranscriptScrollViewer())
            End If
        End Sub

        Private Sub UpdateTranscriptDetachedAnchorOffset(scroller As ScrollViewer)
            If scroller Is Nothing Then
                Return
            End If

            _transcriptDetachedAnchorOffset = Math.Max(0R, scroller.VerticalOffset)
            DebugTranscriptScroll("detached_anchor", $"update={FormatTranscriptScrollMetric(_transcriptDetachedAnchorOffset.GetValueOrDefault())}", scroller)
        End Sub

        Private Function TryRestoreTranscriptDetachedAnchorOffset(scroller As ScrollViewer) As Boolean
            If scroller Is Nothing OrElse
               _transcriptScrollThumbDragActive OrElse
               _transcriptUserScrollInteractionArmed OrElse
               _transcriptScrollProgrammaticMoveInProgress OrElse
               Not _transcriptDetachedAnchorOffset.HasValue OrElse
               _transcriptScrollFollowMode <> TranscriptScrollFollowMode.DetachedByUser Then
                Return False
            End If

            Dim desiredOffset = _transcriptDetachedAnchorOffset.Value
            If Double.IsNaN(desiredOffset) OrElse Double.IsInfinity(desiredOffset) Then
                Return False
            End If

            desiredOffset = Math.Max(0R, Math.Min(desiredOffset, scroller.ScrollableHeight))
            If Math.Abs(scroller.VerticalOffset - desiredOffset) <= 0.5R Then
                Return False
            End If

            DebugTranscriptScroll("detached_anchor", $"restore={FormatTranscriptScrollMetric(desiredOffset)}", scroller)

            _transcriptScrollProgrammaticMoveInProgress = True
            _suppressTranscriptScrollTracking = True
            Try
                scroller.ScrollToVerticalOffset(desiredOffset)
            Finally
                _suppressTranscriptScrollTracking = False
                _transcriptScrollProgrammaticMoveInProgress = False
            End Try

            Return True
        End Function

        Private Sub DebugTranscriptScroll(eventName As String,
                                          Optional details As String = Nothing,
                                          Optional scroller As ScrollViewer = Nothing,
                                          Optional scrollArgs As ScrollChangedEventArgs = Nothing)
            If Not TranscriptScrollDebugInstrumentationEnabled Then
                Return
            End If

            If _viewModel Is Nothing OrElse _viewModel.TranscriptPanel Is Nothing Then
                Return
            End If

            Dim sb As New StringBuilder()
            sb.Append("transcript_scroll ")
            sb.Append("event=").Append(eventName)

            If Not String.IsNullOrWhiteSpace(details) Then
                sb.Append(" details=""").Append(details.Replace("""", "'")).Append(""""c)
            End If

            sb.Append(" follow=").Append(_transcriptScrollFollowMode.ToString())
            sb.Append(" thumbDrag=").Append(_transcriptScrollThumbDragActive)
            sb.Append(" userArmed=").Append(_transcriptUserScrollInteractionArmed)
            sb.Append(" progMove=").Append(_transcriptScrollProgrammaticMoveInProgress)
            sb.Append(" reqPending=").Append(_transcriptScrollRequestPending)
            sb.Append(" reqGen=").Append(_transcriptScrollRequestGeneration)
            sb.Append(" reqPolicy=").Append(_transcriptScrollQueuedPolicy.ToString())
            sb.Append(" reqReasons=").Append(_transcriptScrollQueuedReasons.ToString())
            sb.Append(" reqEpoch=").Append(_transcriptScrollQueuedInteractionEpoch)
            sb.Append(" userEpoch=").Append(_transcriptScrollUserInteractionEpoch)
            sb.Append(" toBottomPending=").Append(_transcriptScrollToBottomPending)
            sb.Append(" toBottomGen=").Append(_transcriptScrollToBottomQueuedGeneration)
            sb.Append(" toBottomPolicy=").Append(_transcriptScrollToBottomQueuedPolicy.ToString())
            sb.Append(" toBottomEpoch=").Append(_transcriptScrollToBottomQueuedInteractionEpoch)

            Dim effectiveScroller = If(scroller, _transcriptScrollViewer)
            If effectiveScroller IsNot Nothing Then
                sb.Append(" offset=").Append(FormatTranscriptScrollMetric(effectiveScroller.VerticalOffset))
                sb.Append(" scrollable=").Append(FormatTranscriptScrollMetric(effectiveScroller.ScrollableHeight))
                sb.Append(" extent=").Append(FormatTranscriptScrollMetric(effectiveScroller.ExtentHeight))
                sb.Append(" viewport=").Append(FormatTranscriptScrollMetric(effectiveScroller.ViewportHeight))
                sb.Append(" atBottom=").Append(IsScrollViewerAtBottomExtreme(effectiveScroller))
            End If

            If scrollArgs IsNot Nothing Then
                sb.Append(" vChange=").Append(FormatTranscriptScrollMetric(scrollArgs.VerticalChange))
                sb.Append(" eChange=").Append(FormatTranscriptScrollMetric(scrollArgs.ExtentHeightChange))
                sb.Append(" vpChange=").Append(FormatTranscriptScrollMetric(scrollArgs.ViewportHeightChange))
            End If

            Dim visibleThreadId = GetVisibleThreadId()
            If Not String.IsNullOrWhiteSpace(visibleThreadId) Then
                sb.Append(" thread=").Append(visibleThreadId)
            End If

            Dim visibleTurnId = GetVisibleTurnId()
            If Not String.IsNullOrWhiteSpace(visibleTurnId) Then
                sb.Append(" turn=").Append(visibleTurnId)
            End If

            AppendProtocol("debug", sb.ToString())
        End Sub

        Private Shared Function FormatTranscriptScrollMetric(value As Double) As String
            If Double.IsNaN(value) OrElse Double.IsInfinity(value) Then
                Return "nan"
            End If

            Return value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Sub CancelPendingTranscriptScrollRequests()
            _transcriptScrollQueuedPolicy = TranscriptScrollRequestPolicy.None
            _transcriptScrollQueuedReasons = TranscriptScrollRequestReason.None
            _transcriptScrollQueuedInteractionEpoch = 0
            _transcriptScrollRequestGeneration += 1
            DebugTranscriptScroll("cancel_pending_requests", "invalidate_generation")
        End Sub

        Private Shared Function ResolveTranscriptScrollRequestPolicy(reason As TranscriptScrollRequestReason,
                                                                    force As Boolean) As TranscriptScrollRequestPolicy
            If force Then
                Return TranscriptScrollRequestPolicy.ForceJump
            End If

            If reason = TranscriptScrollRequestReason.None Then
                Return TranscriptScrollRequestPolicy.None
            End If

            If TranscriptScrollRequestHasReason(reason, TranscriptScrollRequestReason.ThreadSelection) OrElse
               TranscriptScrollRequestHasReason(reason, TranscriptScrollRequestReason.ThreadRebuild) OrElse
               TranscriptScrollRequestHasReason(reason, TranscriptScrollRequestReason.UserMessage) Then
                Return TranscriptScrollRequestPolicy.ForceJump
            End If

            Return TranscriptScrollRequestPolicy.FollowIfPinned
        End Function

        Private Shared Function MergeTranscriptScrollRequestPolicy(currentPolicy As TranscriptScrollRequestPolicy,
                                                                  incomingPolicy As TranscriptScrollRequestPolicy) As TranscriptScrollRequestPolicy
            If currentPolicy = TranscriptScrollRequestPolicy.ForceJump OrElse
               incomingPolicy = TranscriptScrollRequestPolicy.ForceJump Then
                Return TranscriptScrollRequestPolicy.ForceJump
            End If

            If currentPolicy = TranscriptScrollRequestPolicy.FollowIfPinned OrElse
               incomingPolicy = TranscriptScrollRequestPolicy.FollowIfPinned Then
                Return TranscriptScrollRequestPolicy.FollowIfPinned
            End If

            Return TranscriptScrollRequestPolicy.None
        End Function

        Private Shared Function IsRuntimeOnlyTranscriptScrollRequest(reasons As TranscriptScrollRequestReason) As Boolean
            If reasons = TranscriptScrollRequestReason.None Then
                Return False
            End If

            If Not TranscriptScrollRequestHasReason(reasons, TranscriptScrollRequestReason.RuntimeStream) Then
                Return False
            End If

            If TranscriptScrollRequestHasReason(reasons, TranscriptScrollRequestReason.LegacyCallsite) Then
                Return False
            End If

            If TranscriptScrollRequestHasReason(reasons, TranscriptScrollRequestReason.ThreadSelection) OrElse
               TranscriptScrollRequestHasReason(reasons, TranscriptScrollRequestReason.ThreadRebuild) OrElse
               TranscriptScrollRequestHasReason(reasons, TranscriptScrollRequestReason.SystemMessage) OrElse
               TranscriptScrollRequestHasReason(reasons, TranscriptScrollRequestReason.UserMessage) Then
                Return False
            End If

            Return True
        End Function

        Private Shared Function TranscriptScrollRequestHasReason(reasons As TranscriptScrollRequestReason,
                                                                 expected As TranscriptScrollRequestReason) As Boolean
            Return (reasons And expected) = expected
        End Function

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

        Private Shared Function FindVisualDescendantByName(Of T As FrameworkElement)(root As DependencyObject,
                                                                                      elementName As String) As T
            If root Is Nothing OrElse String.IsNullOrWhiteSpace(elementName) Then
                Return Nothing
            End If

            Dim childCount = VisualTreeHelper.GetChildrenCount(root)
            For i = 0 To childCount - 1
                Dim child = VisualTreeHelper.GetChild(root, i)
                Dim typed = TryCast(child, T)
                If typed IsNot Nothing AndAlso StringComparer.Ordinal.Equals(typed.Name, elementName) Then
                    Return typed
                End If

                Dim nested = FindVisualDescendantByName(Of T)(child, elementName)
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
