Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Threading
Imports CodexNativeAgent.AppServer
Imports CodexNativeAgent.Services

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Inherits Window

        Private NotInheritable Class ModelListEntry
            Public Property Id As String = String.Empty
            Public Property DisplayName As String = String.Empty
            Public Property IsDefault As Boolean

            Public Overrides Function ToString() As String
                Dim name = If(String.IsNullOrWhiteSpace(DisplayName), Id, DisplayName)
                If IsDefault Then
                    Return $"{name} ({Id}) [default]"
                End If

                Return $"{name} ({Id})"
            End Function
        End Class

        Private NotInheritable Class ThreadListEntry
            Public Property Id As String = String.Empty
            Public Property Preview As String = String.Empty
            Public Property LastActiveAt As String = String.Empty
            Public Property LastActiveSortTimestamp As Long
            Public Property Cwd As String = String.Empty
            Public Property IsArchived As Boolean

            Public ReadOnly Property ListLeftText As String
                Get
                    Return NormalizePreviewSnippet(Preview)
                End Get
            End Property

            Public ReadOnly Property ListRightText As String
                Get
                    Return FormatCompactAge(LastActiveSortTimestamp)
                End Get
            End Property

            Public ReadOnly Property ListLeftMargin As Thickness
                Get
                    Return New Thickness(18, 0, 0, 0)
                End Get
            End Property

            Public ReadOnly Property ListLeftFontWeight As FontWeight
                Get
                    Return FontWeights.Normal
                End Get
            End Property

            Public Overrides Function ToString() As String
                Dim snippet = ListLeftText
                Dim age = ListRightText
                If String.IsNullOrWhiteSpace(age) Then
                    Return $"    {snippet}"
                End If

                Return $"    {snippet} | {age}"
            End Function

            Private Shared Function NormalizePreviewSnippet(value As String) As String
                Dim text = If(String.IsNullOrWhiteSpace(value), "(untitled)", value)
                text = text.Replace(ControlChars.Cr, " "c).
                            Replace(ControlChars.Lf, " "c).
                            Replace(ControlChars.Tab, " "c).
                            Trim()

                Do While text.Contains("  ", StringComparison.Ordinal)
                    text = text.Replace("  ", " ", StringComparison.Ordinal)
                Loop

                Const maxLength As Integer = 72
                If text.Length > maxLength Then
                    Return text.Substring(0, maxLength - 3) & "..."
                End If

                Return text
            End Function

            Private Shared Function FormatCompactAge(unixMilliseconds As Long) As String
                If unixMilliseconds <= 0 OrElse unixMilliseconds = Long.MinValue Then
                    Return String.Empty
                End If

                Try
                    Dim age = DateTimeOffset.UtcNow - DateTimeOffset.FromUnixTimeMilliseconds(unixMilliseconds)
                    If age < TimeSpan.Zero Then
                        age = TimeSpan.Zero
                    End If

                    If age.TotalMinutes < 1 Then
                        Return "now"
                    End If

                    If age.TotalHours < 1 Then
                        Return $"{Math.Max(1, CInt(Math.Floor(age.TotalMinutes)))}m"
                    End If

                    If age.TotalDays < 1 Then
                        Return $"{Math.Max(1, CInt(Math.Floor(age.TotalHours)))}h"
                    End If

                    If age.TotalDays < 7 Then
                        Return $"{Math.Max(1, CInt(Math.Floor(age.TotalDays)))}d"
                    End If

                    If age.TotalDays < 30 Then
                        Return $"{Math.Max(1, CInt(Math.Floor(age.TotalDays / 7)))}w"
                    End If

                    If age.TotalDays < 365 Then
                        Return $"{Math.Max(1, CInt(Math.Floor(age.TotalDays / 30)))}mo"
                    End If

                    Return $"{Math.Max(1, CInt(Math.Floor(age.TotalDays / 365)))}y"
                Catch
                    Return String.Empty
                End Try
            End Function
        End Class

        Private NotInheritable Class ThreadGroupHeaderEntry
            Public Property GroupKey As String = String.Empty
            Public Property FolderName As String = String.Empty
            Public Property Count As Integer
            Public Property IsExpanded As Boolean

            Public ReadOnly Property ListLeftText As String
                Get
                    Dim folderIcon = Char.ConvertFromUtf32(If(IsExpanded, &H1F4C2, &H1F4C1))
                    Return $"{folderIcon} {FolderName}"
                End Get
            End Property

            Public ReadOnly Property ListRightText As String
                Get
                    Return Count.ToString()
                End Get
            End Property

            Public ReadOnly Property ListLeftMargin As Thickness
                Get
                    Return New Thickness(0)
                End Get
            End Property

            Public ReadOnly Property ListLeftFontWeight As FontWeight
                Get
                    Return FontWeights.SemiBold
                End Get
            End Property

            Public Overrides Function ToString() As String
                Return $"{ListLeftText} ({Count})"
            End Function
        End Class

        Private NotInheritable Class ThreadProjectGroup
            Public Property Key As String = String.Empty
            Public Property HeaderLabel As String = String.Empty
            Public Property LatestActivitySortTimestamp As Long = Long.MinValue
            Public ReadOnly Property Threads As New List(Of ThreadListEntry)()
        End Class

        Private NotInheritable Class PendingUserEcho
            Public Property Text As String = String.Empty
            Public Property AddedUtc As DateTimeOffset
        End Class

        Private NotInheritable Class AppSettings
            Public Property CodexPath As String = String.Empty
            Public Property ServerArgs As String = "app-server"
            Public Property WorkingDir As String = String.Empty
            Public Property WindowsCodexHome As String = String.Empty
            Public Property RememberApiKey As Boolean
            Public Property AutoLoginApiKey As Boolean
            Public Property AutoReconnect As Boolean = True
            Public Property FilterThreadsByWorkingDir As Boolean
            Public Property EncryptedApiKey As String = String.Empty
            Public Property ThemeMode As String = AppAppearanceManager.LightTheme
            Public Property DensityMode As String = AppAppearanceManager.ComfortableDensity
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
        Private ReadOnly _approvalQueue As New Queue(Of PendingApprovalInfo)()
        Private ReadOnly _threadEntries As New List(Of ThreadListEntry)()
        Private ReadOnly _expandedThreadProjectGroups As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _streamingAgentItemIds As New HashSet(Of String)(StringComparer.Ordinal)
        Private ReadOnly _pendingLocalUserEchoes As New Queue(Of PendingUserEcho)()
        Private Shared ReadOnly PendingUserEchoMaxAge As TimeSpan = TimeSpan.FromSeconds(30)

        Private _client As CodexAppServerClient
        Private _currentThreadId As String = String.Empty
        Private _currentTurnId As String = String.Empty
        Private _currentLoginId As String = String.Empty
        Private _disconnecting As Boolean
        Private _threadsLoading As Boolean
        Private _threadLoadError As String = String.Empty
        Private _threadSelectionLoadCts As CancellationTokenSource
        Private _threadSelectionLoadVersion As Integer
        Private _threadContentLoading As Boolean
        Private _suppressThreadSelectionEvents As Boolean
        Private _suppressThreadToolbarMenuEvents As Boolean
        Private _threadContextTarget As ThreadListEntry
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
        Private _activeApproval As PendingApprovalInfo
        Private _currentTheme As String = AppAppearanceManager.LightTheme
        Private _currentDensity As String = AppAppearanceManager.ComfortableDensity
        Private _suppressAppearanceUiChange As Boolean
        Private _settings As New AppSettings()

        Public Sub New()
            InitializeComponent()

            _accountService = New CodexAccountService(Function() CurrentClient())
            _connectionService = New CodexConnectionService()
            _approvalService = New CodexApprovalService()
            _threadService = New CodexThreadService(Function() CurrentClient())
            _turnService = New CodexTurnService(Function() CurrentClient())

            InitializeUiDefaults()
            InitializeEventHandlers()
            InitializeStatusUi()
            InitializeReliabilityLayer()
            InitializeDefaults()
            ShowWorkspaceView()
            RefreshControlStates()
            ShowStatus("Ready.")
        End Sub

        Private Sub InitializeUiDefaults()
            If CmbThreadSort.SelectedIndex < 0 Then
                CmbThreadSort.SelectedIndex = 0
            End If

            If CmbReasoningEffort.SelectedIndex < 0 Then
                CmbReasoningEffort.SelectedIndex = 4
            End If

            If CmbApprovalPolicy.SelectedIndex < 0 Then
                CmbApprovalPolicy.SelectedIndex = 2
            End If

            If CmbSandbox.SelectedIndex < 0 Then
                CmbSandbox.SelectedIndex = 0
            End If

            If CmbDensity.SelectedIndex < 0 Then
                CmbDensity.SelectedIndex = 0
            End If

            If CmbExternalPlanType.Items.Count = 0 Then
                CmbExternalPlanType.Items.Add("")
                CmbExternalPlanType.Items.Add("free")
                CmbExternalPlanType.Items.Add("go")
                CmbExternalPlanType.Items.Add("plus")
                CmbExternalPlanType.Items.Add("pro")
                CmbExternalPlanType.Items.Add("team")
                CmbExternalPlanType.Items.Add("business")
                CmbExternalPlanType.Items.Add("enterprise")
                CmbExternalPlanType.Items.Add("edu")
                CmbExternalPlanType.Items.Add("unknown")
            End If
            CmbExternalPlanType.SelectedIndex = 0

            TxtApproval.Text = "No pending approvals."
            TxtRateLimits.Text = "No rate-limit data loaded yet."
            LblCurrentThread.Text = "New Thread"
            LblCurrentTurn.Text = "Turn: 0"
            LblConnectionState.Text = "Disconnected"
            LblReconnectCountdown.Text = "Reconnect: not scheduled."
            SetTranscriptLoadingState(False)
            InlineApprovalCard.Visibility = Visibility.Collapsed
            UpdateSidebarSelectionState(showSettings:=False)
            SyncAppearanceControls()
            SyncThreadToolbarMenus()
        End Sub

        Private Sub InitializeEventHandlers()
            AddHandler BtnSidebarNewThread.Click, Async Sub(sender, e)
                                                       ShowWorkspaceView()
                                                       Await RunUiActionAsync(AddressOf StartThreadAsync)
                                                   End Sub
            AddHandler BtnSidebarSettings.Click, Sub(sender, e) ShowSettingsView()
            AddHandler BtnSettingsBack.Click, Sub(sender, e) ShowWorkspaceView()
            AddHandler BtnToggleTheme.Click, Sub(sender, e) ToggleTheme()
            AddHandler CmbDensity.SelectionChanged, Sub(sender, e) OnDensitySelectionChanged()

            AddHandler BtnConnect.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf ConnectAsync)
            AddHandler BtnDisconnect.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf DisconnectAsync)
            AddHandler BtnExportDiagnostics.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf ExportDiagnosticsAsync)
            AddHandler BtnReconnectNow.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf ReconnectNowAsync)
            AddHandler ChkAutoReconnect.Checked, Sub(sender, e) SaveSettings()
            AddHandler ChkAutoReconnect.Unchecked, Sub(sender, e) SaveSettings()

            AddHandler BtnAccountRead.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf RefreshAuthenticationGateAsync)
            AddHandler BtnLoginApiKey.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf LoginApiKeyAsync)
            AddHandler BtnLoginChatGpt.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf LoginChatGptAsync)
            AddHandler BtnCancelLogin.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf CancelLoginAsync)
            AddHandler BtnLogout.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf LogoutAsync)
            AddHandler BtnReadRateLimits.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf ReadRateLimitsAsync)
            AddHandler BtnLoginExternalTokens.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf LoginExternalTokensAsync)
            AddHandler ChkRememberApiKey.Checked, Sub(sender, e) SaveSettings()
            AddHandler ChkRememberApiKey.Unchecked, Sub(sender, e) SaveSettings()
            AddHandler ChkAutoLoginApiKey.Checked, Sub(sender, e) SaveSettings()
            AddHandler ChkAutoLoginApiKey.Unchecked, Sub(sender, e) SaveSettings()

            AddHandler ChkShowArchivedThreads.Checked, Async Sub(sender, e) Await RunUiActionAsync(AddressOf RefreshThreadsAsync)
            AddHandler ChkShowArchivedThreads.Unchecked, Async Sub(sender, e) Await RunUiActionAsync(AddressOf RefreshThreadsAsync)
            AddHandler ChkFilterThreadsByWorkingDir.Checked, Async Sub(sender, e) Await RunUiActionAsync(AddressOf RefreshThreadsAsync)
            AddHandler ChkFilterThreadsByWorkingDir.Unchecked, Async Sub(sender, e) Await RunUiActionAsync(AddressOf RefreshThreadsAsync)
            AddHandler TxtThreadSearch.TextChanged, Sub(sender, e) ApplyThreadFiltersAndSort()
            AddHandler CmbThreadSort.SelectionChanged,
                Sub(sender, e)
                    ApplyThreadFiltersAndSort()
                    SyncThreadToolbarMenus()
                End Sub
            AddHandler BtnThreadSortMenu.Click, AddressOf OnThreadSortMenuButtonClick
            AddHandler BtnThreadFilterMenu.Click, AddressOf OnThreadFilterMenuButtonClick
            AddHandler ThreadSortContextMenu.Opened, Sub(sender, e) SyncThreadToolbarMenus()
            AddHandler ThreadFilterContextMenu.Opened, Sub(sender, e) SyncThreadToolbarMenus()
            AddHandler MnuThreadSortNewest.Click, AddressOf OnThreadSortMenuItemClick
            AddHandler MnuThreadSortOldest.Click, AddressOf OnThreadSortMenuItemClick
            AddHandler MnuThreadSortPreviewAz.Click, AddressOf OnThreadSortMenuItemClick
            AddHandler MnuThreadSortPreviewZa.Click, AddressOf OnThreadSortMenuItemClick
            AddHandler MnuThreadFilterArchived.Checked, AddressOf OnThreadFilterMenuItemToggled
            AddHandler MnuThreadFilterArchived.Unchecked, AddressOf OnThreadFilterMenuItemToggled
            AddHandler MnuThreadFilterWorkingDir.Checked, AddressOf OnThreadFilterMenuItemToggled
            AddHandler MnuThreadFilterWorkingDir.Unchecked, AddressOf OnThreadFilterMenuItemToggled
            AddHandler LstThreads.PreviewMouseRightButtonDown, AddressOf OnThreadsPreviewMouseRightButtonDown
            AddHandler LstThreads.ContextMenuOpening, AddressOf OnThreadsContextMenuOpening
            AddHandler ThreadItemContextMenu.Closed, Sub(sender, e) _threadContextTarget = Nothing
            AddHandler MnuThreadSelect.Click, AddressOf OnSelectThreadFromContextMenuClick
            AddHandler MnuThreadRefreshSingle.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf RefreshThreadFromContextMenuAsync)
            AddHandler MnuThreadFork.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf ForkThreadFromContextMenuAsync)
            AddHandler MnuThreadArchive.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf ArchiveThreadFromContextMenuAsync)
            AddHandler MnuThreadUnarchive.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf UnarchiveThreadFromContextMenuAsync)
            AddHandler LstThreads.SelectionChanged,
                Sub(sender, e)
                    If _suppressThreadSelectionEvents Then
                        Return
                    End If

                    Dim selectedHeader = TryCast(LstThreads.SelectedItem, ThreadGroupHeaderEntry)
                    If selectedHeader IsNot Nothing Then
                        _suppressThreadSelectionEvents = True
                        LstThreads.SelectedItem = Nothing
                        _suppressThreadSelectionEvents = False
                        ToggleThreadProjectGroupExpansion(selectedHeader.GroupKey)
                        ApplyThreadFiltersAndSort()
                        Return
                    End If

                    Dim selected = TryCast(LstThreads.SelectedItem, ThreadListEntry)
                    If selected Is Nothing Then
                        CancelActiveThreadSelectionLoad()
                        _threadContentLoading = False
                        SetTranscriptLoadingState(False)
                    Else
                        FireAndForget(AutoLoadThreadSelectionAsync(selected))
                    End If

                    RefreshControlStates()
                End Sub

            AddHandler BtnTurnStart.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf StartTurnAsync)
            AddHandler BtnTurnSteer.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf SteerTurnAsync)
            AddHandler BtnTurnInterrupt.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf InterruptTurnAsync)
            AddHandler BtnClearInput.Click, Sub(sender, e) TxtTurnInput.Clear()
            AddHandler BtnRefreshModels.Click, Async Sub(sender, e) Await RunUiActionAsync(AddressOf RefreshModelsAsync)

            AddHandler BtnApprovalAccept.Click, Async Sub(sender, e) Await RunUiActionAsync(Function() ResolveApprovalAsync("accept"))
            AddHandler BtnApprovalAcceptSession.Click, Async Sub(sender, e) Await RunUiActionAsync(Function() ResolveApprovalAsync("accept_session"))
            AddHandler BtnApprovalDecline.Click, Async Sub(sender, e) Await RunUiActionAsync(Function() ResolveApprovalAsync("decline"))
            AddHandler BtnApprovalCancel.Click, Async Sub(sender, e) Await RunUiActionAsync(Function() ResolveApprovalAsync("cancel"))

            AddHandler BtnQuickOpenVsc.Click, Sub(sender, e) ShowStatus("Open VSC action not implemented yet.")
            AddHandler BtnQuickOpenTerminal.Click, Sub(sender, e) ShowStatus("Open Terminal action not implemented yet.")
        End Sub

        Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
            If _startupConnectAttempted Then
                Return
            End If

            _startupConnectAttempted = True
            FireAndForget(RunUiActionAsync(AddressOf AutoConnectOnStartupAsync))
        End Sub

        Private Sub OnThreadSortMenuButtonClick(sender As Object, e As RoutedEventArgs)
            OpenThreadToolbarMenu(TryCast(sender, Button), ThreadSortContextMenu)
        End Sub

        Private Sub OnThreadFilterMenuButtonClick(sender As Object, e As RoutedEventArgs)
            OpenThreadToolbarMenu(TryCast(sender, Button), ThreadFilterContextMenu)
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
            If MnuThreadSortNewest Is Nothing OrElse
               MnuThreadSortOldest Is Nothing OrElse
               MnuThreadSortPreviewAz Is Nothing OrElse
               MnuThreadSortPreviewZa Is Nothing OrElse
               MnuThreadFilterArchived Is Nothing OrElse
               MnuThreadFilterWorkingDir Is Nothing Then
                Return
            End If

            _suppressThreadToolbarMenuEvents = True
            Try
                Dim sortIndex = Math.Max(0, CmbThreadSort.SelectedIndex)
                MnuThreadSortNewest.IsChecked = (sortIndex = 0)
                MnuThreadSortOldest.IsChecked = (sortIndex = 1)
                MnuThreadSortPreviewAz.IsChecked = (sortIndex = 2)
                MnuThreadSortPreviewZa.IsChecked = (sortIndex = 3)

                MnuThreadFilterArchived.IsChecked = IsChecked(ChkShowArchivedThreads)
                MnuThreadFilterWorkingDir.IsChecked = IsChecked(ChkFilterThreadsByWorkingDir)
            Finally
                _suppressThreadToolbarMenuEvents = False
            End Try
        End Sub

        Private Sub OnThreadSortMenuItemClick(sender As Object, e As RoutedEventArgs)
            If _suppressThreadToolbarMenuEvents Then
                Return
            End If

            Dim targetIndex As Integer = -1
            If ReferenceEquals(sender, MnuThreadSortNewest) Then
                targetIndex = 0
            ElseIf ReferenceEquals(sender, MnuThreadSortOldest) Then
                targetIndex = 1
            ElseIf ReferenceEquals(sender, MnuThreadSortPreviewAz) Then
                targetIndex = 2
            ElseIf ReferenceEquals(sender, MnuThreadSortPreviewZa) Then
                targetIndex = 3
            End If

            If targetIndex < 0 Then
                Return
            End If

            If CmbThreadSort.SelectedIndex <> targetIndex Then
                CmbThreadSort.SelectedIndex = targetIndex
            Else
                SyncThreadToolbarMenus()
            End If

            If ThreadSortContextMenu IsNot Nothing Then
                ThreadSortContextMenu.IsOpen = False
            End If
        End Sub

        Private Sub OnThreadFilterMenuItemToggled(sender As Object, e As RoutedEventArgs)
            If _suppressThreadToolbarMenuEvents Then
                Return
            End If

            Dim archivedChecked = MnuThreadFilterArchived IsNot Nothing AndAlso MnuThreadFilterArchived.IsChecked
            Dim workingDirChecked = MnuThreadFilterWorkingDir IsNot Nothing AndAlso MnuThreadFilterWorkingDir.IsChecked

            Dim changed As Boolean = False
            If ChkShowArchivedThreads IsNot Nothing AndAlso IsChecked(ChkShowArchivedThreads) <> archivedChecked Then
                ChkShowArchivedThreads.IsChecked = archivedChecked
                changed = True
            End If

            If ChkFilterThreadsByWorkingDir IsNot Nothing AndAlso IsChecked(ChkFilterThreadsByWorkingDir) <> workingDirChecked Then
                ChkFilterThreadsByWorkingDir.IsChecked = workingDirChecked
                changed = True
            End If

            If changed Then
                SaveSettings()
            End If

            SyncThreadToolbarMenus()
        End Sub

        Private Sub MainWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles Me.Closing
            _toastTimer.Stop()
            _watchdogTimer.Stop()
            _reconnectUiTimer.Stop()
            CancelReconnect()
            CancelActiveThreadSelectionLoad()
            SaveSettings()

            If _client IsNot Nothing Then
                Try
                    DisconnectAsync().GetAwaiter().GetResult()
                Catch
                End Try
            End If
        End Sub

        Private Sub MainWindow_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Me.PreviewKeyDown
            If Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.Enter Then
                If BtnTurnStart.IsEnabled Then
                    FireAndForget(RunUiActionAsync(AddressOf StartTurnAsync))
                    e.Handled = True
                End If
                Return
            End If

            If e.Key = Key.F5 Then
                If CanRunFullThreadRefresh() Then
                    FireAndForget(RunUiActionAsync(AddressOf RefreshThreadsAsync))
                    e.Handled = True
                End If
                Return
            End If

            If Keyboard.Modifiers = (ModifierKeys.Control Or ModifierKeys.Shift) AndAlso e.Key = Key.N Then
                If BtnSidebarNewThread.IsEnabled Then
                    FireAndForget(RunUiActionAsync(AddressOf StartThreadAsync))
                    e.Handled = True
                End If
                Return
            End If

            If Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.F Then
                ShowThreadsSidebarTab()
                TxtThreadSearch.Focus()
                TxtThreadSearch.SelectAll()
                e.Handled = True
                Return
            End If

            If Keyboard.Modifiers = ModifierKeys.Control AndAlso e.Key = Key.OemComma Then
                ShowControlCenterTab()
                e.Handled = True
            End If
        End Sub

        Private Function CanRunFullThreadRefresh() As Boolean
            Dim connected = _client IsNot Nothing AndAlso _client.IsRunning
            Return connected AndAlso _isAuthenticated AndAlso Not _threadsLoading AndAlso Not _threadContentLoading
        End Function

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
            LblStatus.Text = "Ready."
            LblStatus.Foreground = ResolveBrush("TextPrimaryBrush", Brushes.Black)
        End Sub

        Private Sub ShowStatus(message As String,
                               Optional isError As Boolean = False,
                               Optional displayToast As Boolean = False)
            LblStatus.Text = message
            LblStatus.Foreground = If(isError,
                                      ResolveBrush("DangerBrush", Brushes.DarkRed),
                                      ResolveBrush("TextPrimaryBrush", Brushes.Black))

            If displayToast Then
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

        Private Function ResolveBrush(resourceKey As String, fallback As Brush) As Brush
            If String.IsNullOrWhiteSpace(resourceKey) Then
                Return fallback
            End If

            Dim resolved = TryCast(TryFindResource(resourceKey), Brush)
            Return If(resolved, fallback)
        End Function

        Private Sub OnToastTimerTick(sender As Object, e As EventArgs)
            _toastTimer.Stop()
            ToastOverlay.Visibility = Visibility.Collapsed
        End Sub

        Private Sub UpdateRuntimeFieldState()
            Dim connected = _client IsNot Nothing AndAlso _client.IsRunning
            Dim editable = Not connected

            TxtCodexPath.IsEnabled = editable
            TxtServerArgs.IsEnabled = editable
            TxtWorkingDir.IsEnabled = editable
            TxtWindowsCodexHome.IsEnabled = editable
        End Sub

        Private Sub RefreshControlStates()
            Dim connected = _client IsNot Nothing AndAlso _client.IsRunning
            Dim authenticated = connected AndAlso _isAuthenticated

            BtnConnect.IsEnabled = Not connected
            BtnDisconnect.IsEnabled = connected
            BtnExportDiagnostics.IsEnabled = True
            BtnReconnectNow.IsEnabled = Not connected

            BtnAccountRead.IsEnabled = connected
            BtnLoginApiKey.IsEnabled = connected
            BtnLoginChatGpt.IsEnabled = connected
            BtnCancelLogin.IsEnabled = connected AndAlso Not String.IsNullOrWhiteSpace(_currentLoginId)
            BtnLogout.IsEnabled = authenticated
            BtnReadRateLimits.IsEnabled = authenticated
            BtnLoginExternalTokens.IsEnabled = connected

            BtnRefreshModels.IsEnabled = authenticated

            BtnTurnStart.IsEnabled = authenticated AndAlso Not _threadContentLoading AndAlso Not String.IsNullOrWhiteSpace(_currentThreadId)
            BtnTurnSteer.IsEnabled = authenticated AndAlso Not _threadContentLoading AndAlso Not String.IsNullOrWhiteSpace(_currentThreadId) AndAlso Not String.IsNullOrWhiteSpace(_currentTurnId)
            BtnTurnInterrupt.IsEnabled = authenticated AndAlso Not _threadContentLoading AndAlso Not String.IsNullOrWhiteSpace(_currentThreadId) AndAlso Not String.IsNullOrWhiteSpace(_currentTurnId)
            TxtTurnInput.IsEnabled = authenticated
            BtnSidebarNewThread.IsEnabled = authenticated AndAlso Not _threadContentLoading
            BtnSidebarAutomations.IsEnabled = True
            BtnSidebarSkills.IsEnabled = True
            BtnSidebarSettings.IsEnabled = True
            BtnSettingsBack.IsEnabled = True
            BtnQuickOpenVsc.IsEnabled = True
            BtnQuickOpenTerminal.IsEnabled = True

            Dim hasActiveApproval = _activeApproval IsNot Nothing
            BtnApprovalAccept.IsEnabled = authenticated AndAlso hasActiveApproval
            BtnApprovalAcceptSession.IsEnabled = authenticated AndAlso hasActiveApproval
            BtnApprovalDecline.IsEnabled = authenticated AndAlso hasActiveApproval
            BtnApprovalCancel.IsEnabled = authenticated AndAlso hasActiveApproval
            InlineApprovalCard.Visibility = If(hasActiveApproval, Visibility.Visible, Visibility.Collapsed)

            CmbModel.IsEnabled = authenticated
            CmbReasoningEffort.IsEnabled = authenticated
            CmbApprovalPolicy.IsEnabled = authenticated
            CmbSandbox.IsEnabled = authenticated
            ChkShowArchivedThreads.IsEnabled = authenticated
            ChkFilterThreadsByWorkingDir.IsEnabled = authenticated
            TxtThreadSearch.IsEnabled = authenticated
            CmbThreadSort.IsEnabled = authenticated
            BtnThreadSortMenu.IsEnabled = authenticated
            BtnThreadFilterMenu.IsEnabled = authenticated
            LstThreads.IsEnabled = authenticated AndAlso Not _threadsLoading
            ChkAutoReconnect.IsEnabled = True

            LblWorkspaceHint.Text = BuildWorkspaceHint(connected, authenticated)

            UpdateRuntimeFieldState()
            SyncThreadToolbarMenus()
        End Sub

        Private Function BuildWorkspaceHint(connected As Boolean, authenticated As Boolean) As String
            If Not connected Then
                Return "Connect to Codex App Server from Settings to begin."
            End If

            If Not authenticated Then
                Return "Authentication required: sign in from Settings to unlock threads and turns."
            End If

            If String.IsNullOrWhiteSpace(_currentThreadId) Then
                Return "Start a new thread or select one from the left panel, then send your first instruction."
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
            LeftMainView.Visibility = Visibility.Collapsed
            LeftSettingsView.Visibility = Visibility.Visible
            UpdateSidebarSelectionState(showSettings:=True)
        End Sub

        Private Sub ShowWorkspaceView()
            LeftSettingsView.Visibility = Visibility.Collapsed
            LeftMainView.Visibility = Visibility.Visible
            UpdateSidebarSelectionState(showSettings:=False)
        End Sub

        Private Sub UpdateSidebarSelectionState(showSettings As Boolean)
            BtnSidebarNewThread.Tag = If(showSettings, Nothing, "Active")
            BtnSidebarSettings.Tag = If(showSettings, "Active", Nothing)
            BtnSidebarAutomations.Tag = Nothing
            BtnSidebarSkills.Tag = Nothing
        End Sub

        Private Sub ToggleTheme()
            Dim nextTheme = AppAppearanceManager.ToggleTheme(_currentTheme)
            ApplyAppearance(nextTheme, _currentDensity, persist:=True)
            ShowStatus($"Theme switched to {AppAppearanceManager.DisplayTheme(_currentTheme)}.", displayToast:=True)
        End Sub

        Private Sub OnDensitySelectionChanged()
            If _suppressAppearanceUiChange Then
                Return
            End If

            Dim selectedDensity = AppAppearanceManager.NormalizeDensity(SelectedComboValue(CmbDensity))
            If StringComparer.OrdinalIgnoreCase.Equals(selectedDensity, _currentDensity) Then
                Return
            End If

            ApplyAppearance(_currentTheme, selectedDensity, persist:=True)

            Dim densityLabel = If(StringComparer.OrdinalIgnoreCase.Equals(_currentDensity, AppAppearanceManager.CompactDensity),
                                  "Compact",
                                  "Comfortable")
            ShowStatus($"Density set to {densityLabel}.", displayToast:=True)
        End Sub

        Private Sub ApplyAppearance(themeMode As String, densityMode As String, persist As Boolean)
            _currentTheme = AppAppearanceManager.NormalizeTheme(themeMode)
            _currentDensity = AppAppearanceManager.NormalizeDensity(densityMode)

            AppAppearanceManager.ApplyDensity(_currentDensity)
            AppAppearanceManager.ApplyTheme(_currentTheme)
            SyncAppearanceControls()

            If persist Then
                SaveSettings()
            End If
        End Sub

        Private Sub SyncAppearanceControls()
            _suppressAppearanceUiChange = True
            Try
                Dim compact = StringComparer.OrdinalIgnoreCase.Equals(_currentDensity, AppAppearanceManager.CompactDensity)
                CmbDensity.SelectedIndex = If(compact, 1, 0)
                LblThemeState.Text = $"Current: {AppAppearanceManager.DisplayTheme(_currentTheme)}"
                BtnToggleTheme.Content = AppAppearanceManager.ThemeButtonLabel(_currentTheme)
            Finally
                _suppressAppearanceUiChange = False
            End Try
        End Sub

        Private Function LoadSettings() As AppSettings
            Try
                If Not File.Exists(_settingsFilePath) Then
                    Return New AppSettings()
                End If

                Dim raw = File.ReadAllText(_settingsFilePath)
                Dim loaded = JsonSerializer.Deserialize(Of AppSettings)(raw)
                Return If(loaded, New AppSettings())
            Catch
                Return New AppSettings()
            End Try
        End Function

        Private Sub SaveSettings()
            Try
                CaptureSettingsFromControls()

                If Not _settings.RememberApiKey Then
                    _settings.EncryptedApiKey = String.Empty
                End If

                Dim folder = Path.GetDirectoryName(_settingsFilePath)
                If Not String.IsNullOrWhiteSpace(folder) Then
                    Directory.CreateDirectory(folder)
                End If

                Dim raw = JsonSerializer.Serialize(_settings, _settingsJsonOptions)
                File.WriteAllText(_settingsFilePath, raw)
            Catch
                ' Keep settings failures non-fatal.
            End Try
        End Sub

        Private Sub CaptureSettingsFromControls()
            _settings.CodexPath = TxtCodexPath.Text.Trim()
            _settings.ServerArgs = TxtServerArgs.Text.Trim()
            _settings.WorkingDir = TxtWorkingDir.Text.Trim()
            _settings.WindowsCodexHome = TxtWindowsCodexHome.Text.Trim()
            _settings.RememberApiKey = IsChecked(ChkRememberApiKey)
            _settings.AutoLoginApiKey = IsChecked(ChkAutoLoginApiKey)
            _settings.AutoReconnect = IsChecked(ChkAutoReconnect)
            _settings.FilterThreadsByWorkingDir = IsChecked(ChkFilterThreadsByWorkingDir)
            _settings.ThemeMode = _currentTheme
            _settings.DensityMode = _currentDensity
        End Sub

        Private Sub InitializeDefaults()
            _settings = LoadSettings()
            _settings.ThemeMode = AppAppearanceManager.NormalizeTheme(_settings.ThemeMode)
            _settings.DensityMode = AppAppearanceManager.NormalizeDensity(_settings.DensityMode)

            Dim detectedCodexPath = _connectionService.DetectCodexExecutablePath()

            If String.IsNullOrWhiteSpace(_settings.CodexPath) Then
                _settings.CodexPath = If(String.IsNullOrWhiteSpace(detectedCodexPath), "codex", detectedCodexPath)
            Else
                Dim savedCodexPath = _settings.CodexPath.Trim()
                Dim shouldResolveSavedPath = Not IsPathLike(savedCodexPath) OrElse Not File.Exists(savedCodexPath)

                If shouldResolveSavedPath Then
                    Dim resolvedSavedPath = _connectionService.ResolveWindowsCodexExecutable(savedCodexPath)
                    If Not String.IsNullOrWhiteSpace(resolvedSavedPath) AndAlso
                       ((IsPathLike(resolvedSavedPath) AndAlso File.Exists(resolvedSavedPath)) OrElse Not IsPathLike(resolvedSavedPath)) Then
                        _settings.CodexPath = resolvedSavedPath
                    ElseIf Not String.IsNullOrWhiteSpace(detectedCodexPath) Then
                        _settings.CodexPath = detectedCodexPath
                    End If
                End If
            End If

            If String.IsNullOrWhiteSpace(_settings.ServerArgs) Then
                _settings.ServerArgs = "app-server"
            End If

            If String.IsNullOrWhiteSpace(_settings.WorkingDir) Then
                _settings.WorkingDir = Environment.CurrentDirectory
            End If

            If String.IsNullOrWhiteSpace(_settings.WindowsCodexHome) Then
                _settings.WindowsCodexHome = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex")
            End If

            TxtCodexPath.Text = _settings.CodexPath
            TxtServerArgs.Text = _settings.ServerArgs
            TxtWorkingDir.Text = _settings.WorkingDir
            TxtWindowsCodexHome.Text = _settings.WindowsCodexHome
            ChkRememberApiKey.IsChecked = _settings.RememberApiKey
            ChkAutoLoginApiKey.IsChecked = _settings.AutoLoginApiKey
            ChkAutoReconnect.IsChecked = _settings.AutoReconnect
            ChkFilterThreadsByWorkingDir.IsChecked = _settings.FilterThreadsByWorkingDir
            ApplyAppearance(_settings.ThemeMode, _settings.DensityMode, persist:=False)

            If IsChecked(ChkRememberApiKey) Then
                Dim decryptedApiKey = ReadPersistedApiKey()
                If Not String.IsNullOrWhiteSpace(decryptedApiKey) Then
                    TxtApiKey.Text = decryptedApiKey
                End If
            End If

            UpdateRuntimeFieldState()
            TxtApproval.Text = "No pending approvals."
            TxtRateLimits.Text = "No rate-limit data loaded yet."
            SyncThreadToolbarMenus()
        End Sub

        Private Shared Function IsPathLike(value As String) As Boolean
            If String.IsNullOrWhiteSpace(value) Then
                Return False
            End If

            Return value.Contains(Path.DirectorySeparatorChar) OrElse
                   value.Contains(Path.AltDirectorySeparatorChar) OrElse
                   value.Contains(":"c)
        End Function

        Private Shared Function IsChecked(checkBox As CheckBox) As Boolean
            If checkBox Is Nothing Then
                Return False
            End If

            Return checkBox.IsChecked.HasValue AndAlso checkBox.IsChecked.Value
        End Function

        Private Function CurrentClient() As CodexAppServerClient
            If _client Is Nothing OrElse Not _client.IsRunning Then
                Throw New InvalidOperationException("Not connected to Codex App Server.")
            End If

            Return _client
        End Function

        Private Function EffectiveThreadWorkingDirectory() As String
            Return TxtWorkingDir.Text.Trim()
        End Function

        Private Async Function ConnectAsync() As Task
            Await ConnectCoreAsync(isReconnect:=False, cancellationToken:=CancellationToken.None)
        End Function

        Private Async Function AutoConnectOnStartupAsync() As Task
            If _client IsNot Nothing AndAlso _client.IsRunning Then
                Return
            End If

            ShowStatus("Auto-connecting to Codex App Server...")
            Await ConnectCoreAsync(isReconnect:=False, cancellationToken:=CancellationToken.None)
        End Function

        Private Async Function ConnectCoreAsync(isReconnect As Boolean,
                                                cancellationToken As CancellationToken) As Task
            If _client IsNot Nothing AndAlso _client.IsRunning Then
                Return
            End If

            SaveSettings()

            Dim executable = TxtCodexPath.Text.Trim()
            Dim arguments = TxtServerArgs.Text.Trim()
            Dim workingDir = TxtWorkingDir.Text.Trim()

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
            Dim launchEnvironment = _connectionService.BuildLaunchEnvironment(TxtWindowsCodexHome.Text.Trim())

            TxtCodexPath.Text = launchExecutable

            Dim client = _connectionService.CreateClient()
            AddHandler client.RawMessage, AddressOf ClientOnRawMessage
            AddHandler client.NotificationReceived, AddressOf ClientOnNotification
            AddHandler client.ServerRequestReceived, AddressOf ClientOnServerRequest
            AddHandler client.Disconnected, AddressOf ClientOnDisconnected

            Dim startupError As Exception = Nothing
            Try
                AppendSystemMessage($"Starting '{launchExecutable} {arguments}'...")
                ShowStatus($"Starting '{launchExecutable} {arguments}'...")

                Await _connectionService.StartAndInitializeAsync(client,
                                                                 launchExecutable,
                                                                 arguments,
                                                                 workingDir,
                                                                 launchEnvironment,
                                                                 cancellationToken)

                _client = client
                _connectionExpected = True
                _lastActivityUtc = DateTimeOffset.UtcNow
                _lastWatchdogWarningUtc = DateTimeOffset.MinValue
                _nextReconnectAttemptUtc = Nothing
                LblConnectionState.Text = "Connected"
                UpdateReconnectCountdownUi()

                If isReconnect Then
                    AppendSystemMessage("Reconnected and initialized.")
                    ShowStatus("Reconnected and initialized.", displayToast:=True)
                Else
                    AppendSystemMessage("Connected and initialized.")
                    ShowStatus("Connected and initialized.", displayToast:=True)
                End If

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
            _connectionExpected = False
            CancelReconnect()

            Dim client = _client
            _client = Nothing
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
            LblConnectionState.Text = "Disconnected"
            _currentThreadId = String.Empty
            _currentTurnId = String.Empty
            _currentLoginId = String.Empty
            _isAuthenticated = False
            _authRequiredNoticeShown = False
            _modelsLoadedAtLeastOnce = False
            _threadsLoadedAtLeastOnce = False
            _workspaceBootstrapInProgress = False
            _threadsLoading = False
            _threadLoadError = String.Empty
            _threadContentLoading = False
            _threadEntries.Clear()
            _expandedThreadProjectGroups.Clear()
            _approvalQueue.Clear()
            _activeApproval = Nothing
            _streamingAgentItemIds.Clear()
            _pendingLocalUserEchoes.Clear()
            TxtApproval.Text = "No pending approvals."
            _nextReconnectAttemptUtc = Nothing
            ApplyThreadFiltersAndSort()
            UpdateReconnectCountdownUi()
            UpdateThreadTurnLabels()
            SetTranscriptLoadingState(False)
            RefreshControlStates()
            AppendSystemMessage(message)
            ShowStatus(message, isError:=isError, displayToast:=displayToast)
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
            Dim paramsObject = AsObject(paramsNode)

            Select Case methodName
                Case "thread/started"
                    Dim threadObject = GetPropertyObject(paramsObject, "thread")
                    If threadObject IsNot Nothing Then
                        ApplyCurrentThreadFromThreadObject(threadObject)
                        AppendSystemMessage($"Thread started: {_currentThreadId}")
                    End If

                Case "turn/started"
                    Dim turnObject = GetPropertyObject(paramsObject, "turn")
                    Dim threadId = GetPropertyString(paramsObject, "threadId")
                    If Not String.IsNullOrWhiteSpace(threadId) Then
                        _currentThreadId = threadId
                        MarkThreadLastActive(threadId)
                    End If

                    If turnObject IsNot Nothing Then
                        Dim turnId = GetPropertyString(turnObject, "id")
                        If Not String.IsNullOrWhiteSpace(turnId) Then
                            _currentTurnId = turnId
                        End If
                    End If

                    AppendSystemMessage($"Turn started: {_currentTurnId}")

                Case "turn/completed"
                    Dim turnObject = GetPropertyObject(paramsObject, "turn")
                    Dim threadId = GetPropertyString(paramsObject, "threadId")
                    If Not String.IsNullOrWhiteSpace(threadId) Then
                        MarkThreadLastActive(threadId)
                    ElseIf Not String.IsNullOrWhiteSpace(_currentThreadId) Then
                        MarkThreadLastActive(_currentThreadId)
                    End If

                    Dim completedTurnId = GetPropertyString(turnObject, "id")
                    Dim status = GetPropertyString(turnObject, "status")
                    If StringComparer.Ordinal.Equals(completedTurnId, _currentTurnId) Then
                        _currentTurnId = String.Empty
                    End If

                    _streamingAgentItemIds.Clear()
                    AppendSystemMessage($"Turn completed: {completedTurnId} ({status})")

                Case "item/started"
                    Dim itemObject = GetPropertyObject(paramsObject, "item")
                    If itemObject IsNot Nothing Then
                        Dim itemType = GetPropertyString(itemObject, "type")
                        If StringComparer.Ordinal.Equals(itemType, "agentMessage") Then
                            Dim itemId = GetPropertyString(itemObject, "id")
                            If Not String.IsNullOrWhiteSpace(itemId) Then
                                _streamingAgentItemIds.Add(itemId)
                                TxtTranscript.AppendText($"[{Now:HH:mm:ss}] assistant: ")
                            End If
                        End If
                    End If

                Case "item/agentMessage/delta"
                    Dim delta = GetPropertyString(paramsObject, "delta")
                    Dim itemId = GetPropertyString(paramsObject, "itemId")
                    If String.IsNullOrWhiteSpace(itemId) Then
                        itemId = "live"
                    End If

                    If Not _streamingAgentItemIds.Contains(itemId) Then
                        _streamingAgentItemIds.Add(itemId)
                        TxtTranscript.AppendText($"[{Now:HH:mm:ss}] assistant: ")
                    End If

                    TxtTranscript.AppendText(delta)
                    ScrollTextBoxToBottom(TxtTranscript)

                Case "item/completed"
                    Dim itemObject = GetPropertyObject(paramsObject, "item")
                    If itemObject IsNot Nothing Then
                        Dim itemType = GetPropertyString(itemObject, "type")
                        Dim itemId = GetPropertyString(itemObject, "id")

                        If StringComparer.Ordinal.Equals(itemType, "agentMessage") Then
                            Dim text = GetPropertyString(itemObject, "text")
                            If _streamingAgentItemIds.Contains(itemId) Then
                                TxtTranscript.AppendText(Environment.NewLine & Environment.NewLine)
                                _streamingAgentItemIds.Remove(itemId)
                            ElseIf Not String.IsNullOrWhiteSpace(text) Then
                                AppendTranscript("assistant", text)
                            End If
                        Else
                            RenderItem(itemObject)
                        End If
                    End If

                Case "item/commandExecution/outputDelta"
                    AppendProtocol("cmd", GetPropertyString(paramsObject, "delta"))

                Case "item/fileChange/outputDelta"
                    AppendProtocol("file", GetPropertyString(paramsObject, "delta"))

                Case "item/reasoning/textDelta"
                    AppendProtocol("reason", GetPropertyString(paramsObject, "delta"))

                Case "error"
                    Dim errorObject = GetPropertyObject(paramsObject, "error")
                    Dim message = GetPropertyString(errorObject, "message", "Unknown error")
                    AppendSystemMessage($"Turn error: {message}")

                Case "account/login/completed"
                    Dim success = GetPropertyBoolean(paramsObject, "success", False)
                    Dim loginId = GetPropertyString(paramsObject, "loginId")
                    Dim [error] = GetPropertyString(paramsObject, "error")

                    If StringComparer.Ordinal.Equals(loginId, _currentLoginId) Then
                        _currentLoginId = String.Empty
                    End If

                    If success Then
                        AppendSystemMessage("Account login completed.")
                    Else
                        AppendSystemMessage($"Account login failed: {[error]}")
                    End If

                    RefreshControlStates()
                    FireAndForget(RunUiActionAsync(AddressOf RefreshAuthenticationGateAsync))

                Case "account/updated"
                    FireAndForget(RunUiActionAsync(AddressOf RefreshAuthenticationGateAsync))

                Case "account/rateLimits/updated"
                    If paramsObject IsNot Nothing Then
                        TxtRateLimits.Text = PrettyJson(paramsObject)
                        ShowStatus("Rate limits updated.")
                    End If

                Case "model/rerouted"
                    Dim fromModel = GetPropertyString(paramsObject, "fromModel")
                    Dim toModel = GetPropertyString(paramsObject, "toModel")
                    Dim reason = GetPropertyString(paramsObject, "reason")
                    AppendSystemMessage($"Model rerouted: {fromModel} -> {toModel} ({reason})")

                Case Else
                    ' Keep unsupported notifications in protocol log only.
            End Select

            UpdateThreadTurnLabels()
            RefreshControlStates()
        End Sub

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

            If IsChecked(ChkRememberApiKey) AndAlso Not String.IsNullOrWhiteSpace(apiKey) Then
                Dim plainBytes = Encoding.UTF8.GetBytes(apiKey)
                Dim encryptedBytes = ProtectedData.Protect(plainBytes, Nothing, DataProtectionScope.CurrentUser)
                _settings.EncryptedApiKey = Convert.ToBase64String(encryptedBytes)
            Else
                _settings.EncryptedApiKey = String.Empty
            End If

            SaveSettings()
        End Sub

        Private Async Function TryAutoLoginApiKeyAsync() As Task
            If Not IsChecked(ChkAutoLoginApiKey) Then
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
            Dim wasAuthenticated = _isAuthenticated
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
                _authRequiredNoticeShown = False
                If Not wasAuthenticated OrElse Not _modelsLoadedAtLeastOnce OrElse Not _threadsLoadedAtLeastOnce Then
                    Await InitializeWorkspaceAfterAuthenticationAsync()
                End If

                ShowThreadsSidebarTab()
                Return True
            End If

            ApplyAuthenticationRequiredState(showPrompt:=showAuthPrompt)
            Return False
        End Function

        Private Async Function EnsureAuthenticatedWithRetryAsync(allowAutoLogin As Boolean) As Task(Of Boolean)
            Dim delaysMs As Integer() = {0, 500, 1200, 2500}

            For attempt = 0 To delaysMs.Length - 1
                If delaysMs(attempt) > 0 Then
                    Await Task.Delay(delaysMs(attempt))
                End If

                Dim hasAccount = Await EnsureAuthenticatedAndInitializeWorkspaceAsync(allowAutoLogin:=allowAutoLogin AndAlso attempt = 0,
                                                                                      showAuthPrompt:=attempt = delaysMs.Length - 1)
                If hasAccount Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Async Function RefreshAuthenticationGateAsync() As Task
            Await EnsureAuthenticatedWithRetryAsync(allowAutoLogin:=False)
        End Function

        Private Async Function InitializeWorkspaceAfterAuthenticationAsync() As Task
            If _client Is Nothing OrElse Not _client.IsRunning Then
                Return
            End If

            If Not _isAuthenticated Then
                Return
            End If

            If _workspaceBootstrapInProgress Then
                Return
            End If

            _workspaceBootstrapInProgress = True
            Try
                Await RefreshModelsAsync()
                Await RefreshThreadsAsync()
                ShowStatus("Connected and authenticated.")
            Finally
                _workspaceBootstrapInProgress = False
            End Try
        End Function

        Private Sub ApplyAuthenticationRequiredState(Optional showPrompt As Boolean = False)
            CancelActiveThreadSelectionLoad()
            _isAuthenticated = False
            _currentThreadId = String.Empty
            _currentTurnId = String.Empty
            _threadsLoading = False
            _threadLoadError = String.Empty
            _threadContentLoading = False
            _modelsLoadedAtLeastOnce = False
            _threadsLoadedAtLeastOnce = False
            _workspaceBootstrapInProgress = False
            _threadEntries.Clear()
            _expandedThreadProjectGroups.Clear()
            _streamingAgentItemIds.Clear()
            _pendingLocalUserEchoes.Clear()
            _approvalQueue.Clear()
            _activeApproval = Nothing
            TxtApproval.Text = "No pending approvals."

            LstThreads.Items.Clear()
            CmbModel.Items.Clear()

            UpdateThreadTurnLabels()
            UpdateThreadsStateLabel(0)
            SetTranscriptLoadingState(False)
            RefreshControlStates()

            If showPrompt AndAlso Not _authRequiredNoticeShown Then
                Dim message = "Authentication required. Sign in with ChatGPT or API key to continue."
                AppendSystemMessage(message)
                ShowStatus(message, isError:=True, displayToast:=True)
                _authRequiredNoticeShown = True
            End If

            ShowControlCenterTab()
        End Sub

        Private Async Function RefreshAccountAsync(refreshToken As Boolean) As Task(Of Boolean)
            Dim result = Await _accountService.ReadAccountAsync(refreshToken, CancellationToken.None)

            If result.Account IsNot Nothing Then
                LblAccountState.Text = FormatAccountLabel(result.Account, result.RequiresOpenAiAuth)
            Else
                LblAccountState.Text = If(result.RequiresOpenAiAuth,
                                          "Account: signed out (OpenAI auth required)",
                                          "Account: unknown")
            End If

            Dim hasAccount = result.Account IsNot Nothing
            _isAuthenticated = hasAccount
            If hasAccount Then
                _authRequiredNoticeShown = False
            End If

            RefreshControlStates()
            Return hasAccount
        End Function

        Private Async Function ReadRateLimitsAsync() As Task
            Dim response = Await _accountService.ReadRateLimitsAsync(CancellationToken.None)
            TxtRateLimits.Text = PrettyJson(response)
            AppendSystemMessage("Rate limits updated.")
            ShowStatus("Rate limits updated.")
        End Function

        Private Async Function LoginApiKeyAsync() As Task
            Dim apiKey = TxtApiKey.Text.Trim()
            If String.IsNullOrWhiteSpace(apiKey) Then
                Throw New InvalidOperationException("Enter an OpenAI API key first.")
            End If

            Await _accountService.StartApiKeyLoginAsync(apiKey, CancellationToken.None)
            PersistApiKeyIfNeeded(apiKey)
            AppendSystemMessage("API key login submitted.")
            ShowStatus("API key login submitted.", displayToast:=True)
            Await EnsureAuthenticatedWithRetryAsync(allowAutoLogin:=False)
        End Function

        Private Async Function LoginChatGptAsync() As Task
            Dim result = Await _accountService.StartChatGptLoginAsync(CancellationToken.None)
            _currentLoginId = result.LoginId
            Dim authUrl = result.AuthUrl

            If String.IsNullOrWhiteSpace(_currentLoginId) OrElse String.IsNullOrWhiteSpace(authUrl) Then
                Throw New InvalidOperationException("ChatGPT login did not return a valid auth URL.")
            End If

            Try
                Process.Start(New ProcessStartInfo(authUrl) With {.UseShellExecute = True})
            Catch ex As Exception
                AppendSystemMessage($"Could not open browser automatically: {ex.Message}")
                ShowStatus($"Could not open browser: {ex.Message}", isError:=True)
            End Try

            LblAccountState.Text = $"Account: waiting for browser sign-in (loginId={_currentLoginId})"
            AppendSystemMessage("ChatGPT sign-in started. Finish auth in your browser.")
            ShowStatus("ChatGPT sign-in started.", displayToast:=True)
            RefreshControlStates()
        End Function

        Private Async Function CancelLoginAsync() As Task
            If String.IsNullOrWhiteSpace(_currentLoginId) Then
                Throw New InvalidOperationException("No active ChatGPT login flow to cancel.")
            End If

            Await _accountService.CancelLoginAsync(_currentLoginId, CancellationToken.None)
            AppendSystemMessage($"Canceled login {_currentLoginId}.")
            ShowStatus($"Canceled login {_currentLoginId}.")
            _currentLoginId = String.Empty
            RefreshControlStates()
            Await RefreshAccountAsync(False)
        End Function

        Private Async Function LogoutAsync() As Task
            Await _accountService.LogoutAsync(CancellationToken.None)
            AppendSystemMessage("Logged out.")
            ShowStatus("Logged out.", displayToast:=True)
            _currentLoginId = String.Empty
            RefreshControlStates()
            Await RefreshAccountAsync(False)
            ApplyAuthenticationRequiredState()
        End Function

        Private Async Function LoginExternalTokensAsync() As Task
            Dim token = TxtExternalAccessToken.Text.Trim()
            Dim accountId = TxtExternalAccountId.Text.Trim()

            If String.IsNullOrWhiteSpace(token) Then
                Throw New InvalidOperationException("Enter an external ChatGPT access token.")
            End If

            If String.IsNullOrWhiteSpace(accountId) Then
                Throw New InvalidOperationException("Enter a ChatGPT account/workspace id.")
            End If

            Dim plan = SelectedComboValue(CmbExternalPlanType)
            Await _accountService.StartExternalTokenLoginAsync(token,
                                                               accountId,
                                                               plan,
                                                               CancellationToken.None)

            AppendSystemMessage("External ChatGPT auth tokens applied.")
            ShowStatus("External ChatGPT auth tokens applied.", displayToast:=True)
            Await EnsureAuthenticatedWithRetryAsync(allowAutoLogin:=False)
        End Function

        Private Async Function HandleChatgptTokenRefreshAsync(request As RpcServerRequest) As Task
            Dim token = TxtExternalAccessToken.Text.Trim()
            Dim accountId = TxtExternalAccountId.Text.Trim()

            If String.IsNullOrWhiteSpace(token) OrElse String.IsNullOrWhiteSpace(accountId) Then
                Await CurrentClient().SendErrorAsync(request.Id,
                                                     -32001,
                                                     "ChatGPT auth token refresh requested, but external token/account id are not configured in the UI.")
                AppendSystemMessage("Could not refresh external ChatGPT token: missing token/account id.")
                ShowStatus("Could not refresh external token: missing token/account id.", isError:=True)
                Return
            End If

            Dim response As New JsonObject()
            response("accessToken") = token
            response("chatgptAccountId") = accountId

            Dim planType = SelectedComboValue(CmbExternalPlanType)
            If Not String.IsNullOrWhiteSpace(planType) Then
                response("chatgptPlanType") = planType
            End If

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

        Private Async Function RefreshModelsAsync() As Task
            Dim previousModelId = SelectedModelId()
            Dim models = Await _threadService.ListModelsAsync(CancellationToken.None)

            CmbModel.Items.Clear()
            For Each model In models
                CmbModel.Items.Add(New ModelListEntry() With {
                    .Id = model.Id,
                    .DisplayName = model.DisplayName,
                    .IsDefault = model.IsDefault
                })
            Next

            Dim selectedIndex = -1
            For i = 0 To CmbModel.Items.Count - 1
                Dim item = TryCast(CmbModel.Items(i), ModelListEntry)
                If item Is Nothing Then
                    Continue For
                End If

                If StringComparer.Ordinal.Equals(item.Id, previousModelId) Then
                    selectedIndex = i
                    Exit For
                End If

                If selectedIndex = -1 AndAlso item.IsDefault Then
                    selectedIndex = i
                End If
            Next

            If selectedIndex >= 0 Then
                CmbModel.SelectedIndex = selectedIndex
            End If

            _modelsLoadedAtLeastOnce = True
            ShowStatus($"Loaded {CmbModel.Items.Count} model(s).")
            AppendSystemMessage($"Loaded {CmbModel.Items.Count} models.")
        End Function

        Private Async Function StartThreadAsync() As Task
            CancelActiveThreadSelectionLoad()
            _threadContentLoading = False
            SetTranscriptLoadingState(False)

            Dim options = BuildThreadRequestOptions(includeModel:=True)
            Dim threadObject = Await _threadService.StartThreadAsync(options, CancellationToken.None)

            ApplyCurrentThreadFromThreadObject(threadObject)
            TxtTranscript.Clear()
            RenderThreadObject(threadObject)

            AppendSystemMessage($"Started thread {_currentThreadId}.")
            ShowStatus($"Started thread {_currentThreadId}.", displayToast:=True)
            Await RefreshThreadsAsync()
        End Function

        Private Async Function RefreshThreadsAsync() As Task
            If _client Is Nothing OrElse Not _client.IsRunning Then
                Return
            End If

            _threadsLoading = True
            _threadLoadError = String.Empty
            UpdateThreadsStateLabel(VisibleThreadCount())
            RefreshControlStates()
            ShowStatus("Loading threads...")

            Try
                Dim useWorkingDirFilter = IsChecked(ChkFilterThreadsByWorkingDir)
                Dim cwd = If(useWorkingDirFilter, EffectiveThreadWorkingDirectory(), String.Empty)
                Dim summaries = Await _threadService.ListThreadsAsync(IsChecked(ChkShowArchivedThreads),
                                                                      cwd,
                                                                      CancellationToken.None)

                _threadEntries.Clear()
                For Each summary In summaries
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

                _threadsLoadedAtLeastOnce = True
                ApplyThreadFiltersAndSort()
                AppendSystemMessage($"Loaded {_threadEntries.Count} thread(s).")
                ShowStatus($"Loaded {_threadEntries.Count} thread(s).")
            Catch ex As Exception
                _threadLoadError = ex.Message
                UpdateThreadsStateLabel(VisibleThreadCount())
                ShowStatus($"Could not load threads: {ex.Message}", isError:=True, displayToast:=True)
                Throw
            Finally
                _threadsLoading = False
                UpdateThreadsStateLabel(VisibleThreadCount())
                RefreshControlStates()
            End Try
        End Function

        Private Async Function AutoLoadThreadSelectionAsync(selected As ThreadListEntry,
                                                            Optional forceReload As Boolean = False) As Task
            If selected Is Nothing OrElse String.IsNullOrWhiteSpace(selected.Id) Then
                Return
            End If

            If _threadsLoading OrElse _client Is Nothing OrElse Not _client.IsRunning OrElse Not _isAuthenticated Then
                Return
            End If

            Dim selectedThreadId = selected.Id.Trim()
            If Not forceReload AndAlso
               StringComparer.Ordinal.Equals(selectedThreadId, _currentThreadId) AndAlso
               Not _threadContentLoading Then
                Return
            End If

            Dim loadVersion = Interlocked.Increment(_threadSelectionLoadVersion)
            CancelActiveThreadSelectionLoad()

            Dim threadLoadCts As New CancellationTokenSource()
            _threadSelectionLoadCts = threadLoadCts
            Dim cancellationToken = threadLoadCts.Token

            _threadContentLoading = True
            SetTranscriptLoadingState(True, "Loading selected thread...")
            RefreshControlStates()
            ShowStatus("Loading selected thread...")

            Try
                Dim threadObject = Await _threadService.ReadThreadAsync(selectedThreadId,
                                                                        includeTurns:=True,
                                                                        cancellationToken:=cancellationToken).ConfigureAwait(False)
                cancellationToken.ThrowIfCancellationRequested()

                Dim hasTurns = ThreadObjectHasTurns(threadObject)
                Dim transcriptSnapshot = Await Task.Run(Function() BuildThreadTranscriptSnapshot(threadObject), cancellationToken).ConfigureAwait(False)
                cancellationToken.ThrowIfCancellationRequested()

                Await RunOnUiAsync(
                    Function()
                        If cancellationToken.IsCancellationRequested OrElse loadVersion <> _threadSelectionLoadVersion Then
                            Return Task.CompletedTask
                        End If

                        Dim selectedNow = TryCast(LstThreads.SelectedItem, ThreadListEntry)
                        If selectedNow Is Nothing OrElse Not StringComparer.Ordinal.Equals(selectedNow.Id, selectedThreadId) Then
                            Return Task.CompletedTask
                        End If

                        ApplyCurrentThreadFromThreadObject(threadObject)
                        ApplyThreadTranscriptSnapshot(transcriptSnapshot, hasTurns)
                        AppendSystemMessage($"Loaded thread {_currentThreadId} from history.")
                        ShowStatus($"Loaded thread {_currentThreadId}.")
                        Return Task.CompletedTask
                    End Function).ConfigureAwait(False)
            Catch ex As OperationCanceledException
            Catch ex As Exception
                RunOnUi(
                    Sub()
                        If loadVersion = _threadSelectionLoadVersion Then
                            ShowStatus($"Could not load thread {selectedThreadId}: {ex.Message}", isError:=True, displayToast:=True)
                            AppendTranscript("system", $"Could not load thread {selectedThreadId}: {ex.Message}")
                        End If
                    End Sub)
            Finally
                RunOnUi(
                    Sub()
                        If loadVersion = _threadSelectionLoadVersion Then
                            _threadContentLoading = False
                            SetTranscriptLoadingState(False)
                            RefreshControlStates()
                        End If
                    End Sub)

                If _threadSelectionLoadCts Is threadLoadCts Then
                    _threadSelectionLoadCts = Nothing
                End If

                threadLoadCts.Dispose()
            End Try
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
            TranscriptLoadingOverlay.Visibility = If(isLoading, Visibility.Visible, Visibility.Collapsed)
            LblTranscriptLoading.Text = If(isLoading, loadingText, "Loading thread...")
        End Sub

        Private Sub OnThreadsPreviewMouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)
            _threadContextTarget = Nothing

            Dim source = TryCast(e.OriginalSource, DependencyObject)
            Dim container = FindVisualAncestor(Of ListBoxItem)(source)
            If container Is Nothing Then
                e.Handled = True
                Return
            End If

            Dim entry = TryCast(container.DataContext, ThreadListEntry)
            If entry Is Nothing Then
                e.Handled = True
                Return
            End If

            _threadContextTarget = entry
            e.Handled = True
            PrepareThreadContextMenu(entry)
            ThreadItemContextMenu.PlacementTarget = container
            ThreadItemContextMenu.IsOpen = True
        End Sub

        Private Sub OnThreadsContextMenuOpening(sender As Object, e As ContextMenuEventArgs)
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

            MnuThreadSelect.IsEnabled = True

            Dim canRunThreadAction = _isAuthenticated AndAlso Not _threadsLoading AndAlso Not _threadContentLoading
            MnuThreadRefreshSingle.IsEnabled = canRunThreadAction
            MnuThreadFork.IsEnabled = canRunThreadAction

            Dim archived = target.IsArchived
            MnuThreadArchive.Visibility = If(archived, Visibility.Collapsed, Visibility.Visible)
            MnuThreadUnarchive.Visibility = If(archived, Visibility.Visible, Visibility.Collapsed)
            MnuThreadArchive.IsEnabled = canRunThreadAction AndAlso Not archived
            MnuThreadUnarchive.IsEnabled = canRunThreadAction AndAlso archived
        End Sub

        Private Sub OnSelectThreadFromContextMenuClick(sender As Object, e As RoutedEventArgs)
            Dim target = ResolveContextThreadEntry()
            If target Is Nothing Then
                Return
            End If

            SelectThreadEntry(target, suppressAutoLoad:=False)
        End Sub

        Private Async Function RefreshThreadFromContextMenuAsync() As Task
            Dim target = ResolveContextThreadEntry()
            If target Is Nothing Then
                Throw New InvalidOperationException("Select a thread first.")
            End If

            SelectThreadEntry(target, suppressAutoLoad:=True)
            Await AutoLoadThreadSelectionAsync(target, forceReload:=True)
        End Function

        Private Async Function ForkThreadFromContextMenuAsync() As Task
            Dim target = ResolveContextThreadEntry()
            If target Is Nothing Then
                Throw New InvalidOperationException("Select a thread first.")
            End If

            Await ForkThreadAsync(target)
        End Function

        Private Async Function ArchiveThreadFromContextMenuAsync() As Task
            Dim target = ResolveContextThreadEntry()
            If target Is Nothing Then
                Throw New InvalidOperationException("Select a thread first.")
            End If

            Await ArchiveThreadAsync(target)
        End Function

        Private Async Function UnarchiveThreadFromContextMenuAsync() As Task
            Dim target = ResolveContextThreadEntry()
            If target Is Nothing Then
                Throw New InvalidOperationException("Select a thread first.")
            End If

            Await UnarchiveThreadAsync(target)
        End Function

        Private Function ResolveContextThreadEntry() As ThreadListEntry
            If _threadContextTarget IsNot Nothing Then
                Return _threadContextTarget
            End If

            Return TryCast(LstThreads.SelectedItem, ThreadListEntry)
        End Function

        Private Sub SelectThreadEntry(entry As ThreadListEntry, suppressAutoLoad As Boolean)
            If entry Is Nothing Then
                Return
            End If

            If suppressAutoLoad Then
                _suppressThreadSelectionEvents = True
            End If

            Try
                LstThreads.SelectedItem = entry
            Finally
                If suppressAutoLoad Then
                    _suppressThreadSelectionEvents = False
                End If
            End Try
        End Sub

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
            Dim options = BuildThreadRequestOptions(includeModel:=False)
            Dim threadObject = Await _threadService.ForkThreadAsync(selected.Id, options, CancellationToken.None)

            ApplyCurrentThreadFromThreadObject(threadObject)
            RenderThreadObject(threadObject)
            AppendSystemMessage($"Forked into new thread {_currentThreadId}.")
            ShowStatus($"Forked thread {_currentThreadId}.", displayToast:=True)
            Await RefreshThreadsAsync()
        End Function

        Private Async Function ArchiveThreadAsync(selected As ThreadListEntry) As Task
            Await _threadService.ArchiveThreadAsync(selected.Id, CancellationToken.None)
            AppendSystemMessage($"Archived thread {selected.Id}.")
            ShowStatus($"Archived thread {selected.Id}.")
            Await RefreshThreadsAsync()
        End Function

        Private Async Function UnarchiveThreadAsync(selected As ThreadListEntry) As Task
            Await _threadService.UnarchiveThreadAsync(selected.Id, CancellationToken.None)
            AppendSystemMessage($"Unarchived thread {selected.Id}.")
            ShowStatus($"Unarchived thread {selected.Id}.")
            Await RefreshThreadsAsync()
        End Function

        Private Sub ApplyThreadFiltersAndSort()
            Dim searchText = TxtThreadSearch.Text.Trim()
            Dim forceExpandMatchingGroups = Not String.IsNullOrWhiteSpace(searchText)
            Dim filtered As New List(Of ThreadListEntry)()

            For Each entry In _threadEntries
                If MatchesThreadSearch(entry, searchText) Then
                    filtered.Add(entry)
                End If
            Next

            filtered.Sort(AddressOf CompareThreadEntries)

            Dim selectedThreadId As String = _currentThreadId
            Dim selectedEntry = TryCast(LstThreads.SelectedItem, ThreadListEntry)
            If selectedEntry IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(selectedEntry.Id) Then
                selectedThreadId = selectedEntry.Id
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
                LstThreads.Items.Clear()
                For Each group In orderedGroups
                    Dim isExpanded = forceExpandMatchingGroups OrElse _expandedThreadProjectGroups.Contains(group.Key)
                    LstThreads.Items.Add(New ThreadGroupHeaderEntry() With {
                        .GroupKey = group.Key,
                        .FolderName = group.HeaderLabel,
                        .Count = group.Threads.Count,
                        .IsExpanded = isExpanded
                    })

                    If isExpanded Then
                        For Each entry In group.Threads
                            LstThreads.Items.Add(entry)
                        Next
                    End If
                Next

                If Not String.IsNullOrWhiteSpace(selectedThreadId) Then
                    For i = 0 To LstThreads.Items.Count - 1
                        Dim entry = TryCast(LstThreads.Items(i), ThreadListEntry)
                        If entry IsNot Nothing AndAlso StringComparer.Ordinal.Equals(entry.Id, selectedThreadId) Then
                            LstThreads.SelectedIndex = i
                            Exit For
                        End If
                    Next
                End If
            Finally
                _suppressThreadSelectionEvents = False
            End Try

            If TryCast(LstThreads.SelectedItem, ThreadListEntry) Is Nothing Then
                CancelActiveThreadSelectionLoad()
                _threadContentLoading = False
                SetTranscriptLoadingState(False)
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
            Select Case CmbThreadSort.SelectedIndex
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
            Select Case CmbThreadSort.SelectedIndex
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

        Private Function VisibleThreadCount() As Integer
            Dim count = 0
            For Each item As Object In LstThreads.Items
                If TypeOf item Is ThreadListEntry Then
                    count += 1
                End If
            Next

            Return count
        End Function

        Private Sub UpdateThreadsStateLabel(displayCount As Integer)
            Dim connected = _client IsNot Nothing AndAlso _client.IsRunning

            If Not connected Then
                LblThreadsState.Text = "Connect to Codex App Server to load threads."
                Return
            End If

            If Not _isAuthenticated Then
                LblThreadsState.Text = "Authentication required. Sign in to start or load threads."
                Return
            End If

            If _threadsLoading Then
                LblThreadsState.Text = "Loading threads..."
                Return
            End If

            If Not String.IsNullOrWhiteSpace(_threadLoadError) Then
                LblThreadsState.Text = $"Error loading threads: {_threadLoadError}"
                Return
            End If

            If _threadEntries.Count = 0 Then
                LblThreadsState.Text = "No threads found. Start a new thread to begin."
                Return
            End If

            If displayCount = 0 Then
                Dim hasProjectHeaders = False
                For Each item As Object In LstThreads.Items
                    If TypeOf item Is ThreadGroupHeaderEntry Then
                        hasProjectHeaders = True
                        Exit For
                    End If
                Next

                If hasProjectHeaders Then
                    LblThreadsState.Text = "All project folders are collapsed. Expand a folder to view threads."
                Else
                    LblThreadsState.Text = "No threads match the current search/filter."
                End If
                Return
            End If

            LblThreadsState.Text = $"Showing {displayCount} of {_threadEntries.Count} thread(s)."
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

        Private Function BuildThreadRequestOptions(includeModel As Boolean) As ThreadRequestOptions
            Dim options As New ThreadRequestOptions() With {
                .ApprovalPolicy = SelectedComboValue(CmbApprovalPolicy),
                .Sandbox = SelectedComboValue(CmbSandbox),
                .Cwd = EffectiveThreadWorkingDirectory()
            }

            If includeModel Then
                options.Model = SelectedModelId()
            End If

            Return options
        End Function

        Private Sub RenderThreadObject(threadObject As JsonObject)
            TxtTranscript.Clear()

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

            ScrollTextBoxToBottom(TxtTranscript)
        End Sub

        Private Sub ApplyThreadTranscriptSnapshot(transcriptSnapshot As String, hasTurns As Boolean)
            TxtTranscript.Clear()
            If Not hasTurns Then
                AppendSystemMessage("No historical turns loaded for this thread.")
                Return
            End If

            TxtTranscript.Text = transcriptSnapshot
            TrimLogIfNeeded(TxtTranscript)
            ScrollTextBoxToBottom(TxtTranscript)
        End Sub

        Private Shared Function ThreadObjectHasTurns(threadObject As JsonObject) As Boolean
            Dim turns = GetPropertyArray(threadObject, "turns")
            Return turns IsNot Nothing AndAlso turns.Count > 0
        End Function

        Private Shared Function BuildThreadTranscriptSnapshot(threadObject As JsonObject) As String
            Dim turns = GetPropertyArray(threadObject, "turns")
            If turns Is Nothing OrElse turns.Count = 0 Then
                Return String.Empty
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
                        AppendSnapshotItem(builder, itemObject)
                    End If
                Next
            Next

            Return builder.ToString().TrimEnd()
        End Function

        Private Shared Sub AppendSnapshotItem(builder As StringBuilder, itemObject As JsonObject)
            Dim itemType = GetPropertyString(itemObject, "type")

            Select Case itemType
                Case "userMessage"
                    Dim content = GetPropertyArray(itemObject, "content")
                    AppendSnapshotLine(builder, "user", FlattenUserInput(content))

                Case "agentMessage"
                    AppendSnapshotLine(builder, "assistant", GetPropertyString(itemObject, "text"))

                Case "plan"
                    AppendSnapshotLine(builder, "plan", GetPropertyString(itemObject, "text"))

                Case "reasoning"
                    AppendSnapshotLine(builder, "reasoning", GetPropertyString(itemObject, "text", "[reasoning item]"))

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

                Case "fileChange"
                    Dim status = GetPropertyString(itemObject, "status")
                    Dim changes = GetPropertyArray(itemObject, "changes")
                    Dim count = If(changes Is Nothing, 0, changes.Count)
                    AppendSnapshotLine(builder, "fileChange", $"{count} change(s), status={status}")

                Case Else
                    Dim itemId = GetPropertyString(itemObject, "id")
                    AppendSnapshotLine(builder, "item", $"{itemType} ({itemId})")
            End Select
        End Sub

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

        Private Sub ApplyCurrentThreadFromThreadObject(threadObject As JsonObject)
            Dim threadId = GetPropertyString(threadObject, "id")
            If Not String.IsNullOrWhiteSpace(threadId) Then
                _currentThreadId = threadId
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
            Dim selected = TryCast(LstThreads.SelectedItem, ThreadListEntry)
            If selected Is Nothing OrElse String.IsNullOrWhiteSpace(selected.Id) Then
                Throw New InvalidOperationException("Select a thread first.")
            End If

            Return selected
        End Function

        Private Function SelectedModelId() As String
            Dim selected = TryCast(CmbModel.SelectedItem, ModelListEntry)
            If selected Is Nothing Then
                Return String.Empty
            End If

            Return selected.Id
        End Function

        Private Shared Function SelectedComboValue(comboBox As ComboBox) As String
            If comboBox Is Nothing OrElse comboBox.SelectedItem Is Nothing Then
                Return String.Empty
            End If

            Dim comboItem = TryCast(comboBox.SelectedItem, ComboBoxItem)
            If comboItem IsNot Nothing Then
                Return If(comboItem.Content, String.Empty).ToString().Trim()
            End If

            Return comboBox.SelectedItem.ToString().Trim()
        End Function

        Private Async Function StartTurnAsync() As Task
            EnsureThreadSelected()

            Dim inputText = TxtTurnInput.Text.Trim()
            If String.IsNullOrWhiteSpace(inputText) Then
                Throw New InvalidOperationException("Enter turn input before sending.")
            End If

            AppendTranscript("user", inputText)
            TrackPendingUserEcho(inputText)

            Dim result = Await _turnService.StartTurnAsync(_currentThreadId,
                                                           inputText,
                                                           SelectedModelId(),
                                                           SelectedComboValue(CmbReasoningEffort),
                                                           SelectedComboValue(CmbApprovalPolicy),
                                                           CancellationToken.None)

            If Not String.IsNullOrWhiteSpace(result.TurnId) Then
                _currentTurnId = result.TurnId
            End If

            MarkThreadLastActive(_currentThreadId)
            UpdateThreadTurnLabels()
            RefreshControlStates()
            TxtTurnInput.Clear()
            ShowStatus($"Turn started: {_currentTurnId}")
        End Function

        Private Async Function SteerTurnAsync() As Task
            EnsureThreadSelected()

            If String.IsNullOrWhiteSpace(_currentTurnId) Then
                Throw New InvalidOperationException("No active turn to steer.")
            End If

            Dim steerText = TxtTurnInput.Text.Trim()
            If String.IsNullOrWhiteSpace(steerText) Then
                Throw New InvalidOperationException("Enter steer input before sending.")
            End If

            AppendTranscript("user (steer)", steerText)
            TrackPendingUserEcho(steerText)

            Dim returnedTurnId = Await _turnService.SteerTurnAsync(_currentThreadId,
                                                                   _currentTurnId,
                                                                   steerText,
                                                                   CancellationToken.None)

            If Not String.IsNullOrWhiteSpace(returnedTurnId) Then
                _currentTurnId = returnedTurnId
            End If

            MarkThreadLastActive(_currentThreadId)
            UpdateThreadTurnLabels()
            RefreshControlStates()
            TxtTurnInput.Clear()
            ShowStatus($"Turn steered: {_currentTurnId}")
        End Function

        Private Async Function InterruptTurnAsync() As Task
            EnsureThreadSelected()

            If String.IsNullOrWhiteSpace(_currentTurnId) Then
                Throw New InvalidOperationException("No active turn to interrupt.")
            End If

            Await _turnService.InterruptTurnAsync(_currentThreadId,
                                                  _currentTurnId,
                                                  CancellationToken.None)

            AppendSystemMessage($"Interrupt requested for turn {_currentTurnId}.")
            ShowStatus($"Interrupt requested for turn {_currentTurnId}.", displayToast:=True)
        End Function

        Private Sub AppendTranscript(role As String, text As String)
            If String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            TxtTranscript.AppendText($"[{Now:HH:mm:ss}] {role}: {text}{Environment.NewLine}{Environment.NewLine}")
            ScrollTextBoxToBottom(TxtTranscript)
            TrimLogIfNeeded(TxtTranscript)
        End Sub

        Private Sub AppendSystemMessage(message As String)
            AppendTranscript("system", message)
            ShowStatus(message)
        End Sub

        Private Sub AppendProtocol(direction As String, payload As String)
            Dim safePayload = If(payload, String.Empty)
            TxtProtocol.AppendText($"[{Now:HH:mm:ss}] {direction}: {safePayload}{Environment.NewLine}")
            ScrollTextBoxToBottom(TxtProtocol)
            TrimLogIfNeeded(TxtProtocol)
        End Sub

        Private Sub RenderItem(itemObject As JsonObject)
            Dim itemType = GetPropertyString(itemObject, "type")

            Select Case itemType
                Case "userMessage"
                    Dim content = GetPropertyArray(itemObject, "content")
                    Dim userText = FlattenUserInput(content)
                    If Not ShouldSuppressUserEcho(userText) Then
                        AppendTranscript("user", userText)
                    End If

                Case "agentMessage"
                    AppendTranscript("assistant", GetPropertyString(itemObject, "text"))

                Case "plan"
                    AppendTranscript("plan", GetPropertyString(itemObject, "text"))

                Case "reasoning"
                    AppendTranscript("reasoning", GetPropertyString(itemObject, "text", "[reasoning item]"))

                Case "commandExecution"
                    Dim command = GetPropertyString(itemObject, "command")
                    Dim status = GetPropertyString(itemObject, "status")
                    Dim output = GetPropertyString(itemObject, "aggregatedOutput")
                    Dim summary As New StringBuilder()
                    summary.AppendLine($"Command ({status}): {command}")
                    If Not String.IsNullOrWhiteSpace(output) Then
                        summary.AppendLine(output)
                    End If
                    AppendTranscript("command", summary.ToString().TrimEnd())

                Case "fileChange"
                    Dim status = GetPropertyString(itemObject, "status")
                    Dim changes = GetPropertyArray(itemObject, "changes")
                    Dim count = If(changes Is Nothing, 0, changes.Count)
                    AppendTranscript("fileChange", $"{count} change(s), status={status}")

                Case Else
                    Dim itemId = GetPropertyString(itemObject, "id")
                    AppendTranscript("item", $"{itemType} ({itemId})")
            End Select
        End Sub

        Private Sub UpdateThreadTurnLabels()
            LblCurrentThread.Text = If(String.IsNullOrWhiteSpace(_currentThreadId),
                                       "New Thread",
                                       _currentThreadId)

            LblCurrentTurn.Text = If(String.IsNullOrWhiteSpace(_currentTurnId),
                                     "Turn: 0",
                                     $"Turn: {_currentTurnId}")
        End Sub

        Private Shared Function FlattenUserInput(content As JsonArray) As String
            If content Is Nothing OrElse content.Count = 0 Then
                Return String.Empty
            End If

            Dim parts As New List(Of String)()

            For Each entryNode In content
                Dim entryObject = AsObject(entryNode)
                If entryObject Is Nothing Then
                    Continue For
                End If

                Dim kind = GetPropertyString(entryObject, "type")
                Select Case kind
                    Case "text"
                        Dim value = GetPropertyString(entryObject, "text")
                        If Not String.IsNullOrWhiteSpace(value) Then
                            parts.Add(value)
                        End If
                    Case "image"
                        parts.Add($"[image] {GetPropertyString(entryObject, "url")}")
                    Case "localImage"
                        parts.Add($"[localImage] {GetPropertyString(entryObject, "path")}")
                    Case "mention"
                        parts.Add($"[mention] {GetPropertyString(entryObject, "name")}")
                    Case "skill"
                        parts.Add($"[skill] {GetPropertyString(entryObject, "name")}")
                End Select
            Next

            Return String.Join(Environment.NewLine, parts)
        End Function

        Private Sub TrackPendingUserEcho(text As String)
            Dim normalized = NormalizeUserEchoText(text)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return
            End If

            _pendingLocalUserEchoes.Enqueue(New PendingUserEcho With {
                .Text = normalized,
                .AddedUtc = DateTimeOffset.UtcNow
            })

            PrunePendingUserEchoes(DateTimeOffset.UtcNow)
        End Sub

        Private Function ShouldSuppressUserEcho(text As String) As Boolean
            Dim normalized = NormalizeUserEchoText(text)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return False
            End If

            PrunePendingUserEchoes(DateTimeOffset.UtcNow)

            If _pendingLocalUserEchoes.Count = 0 Then
                Return False
            End If

            If StringComparer.Ordinal.Equals(_pendingLocalUserEchoes.Peek().Text, normalized) Then
                _pendingLocalUserEchoes.Dequeue()
                Return True
            End If

            Dim size = _pendingLocalUserEchoes.Count
            Dim matched = False
            For i = 0 To size - 1
                Dim candidate = _pendingLocalUserEchoes.Dequeue()
                If Not matched AndAlso StringComparer.Ordinal.Equals(candidate.Text, normalized) Then
                    matched = True
                    Continue For
                End If

                _pendingLocalUserEchoes.Enqueue(candidate)
            Next

            Return matched
        End Function

        Private Sub PrunePendingUserEchoes(nowUtc As DateTimeOffset)
            While _pendingLocalUserEchoes.Count > 16
                _pendingLocalUserEchoes.Dequeue()
            End While

            While _pendingLocalUserEchoes.Count > 0
                Dim pending = _pendingLocalUserEchoes.Peek()
                If nowUtc - pending.AddedUtc <= PendingUserEchoMaxAge Then
                    Exit While
                End If

                _pendingLocalUserEchoes.Dequeue()
            End While
        End Sub

        Private Shared Function NormalizeUserEchoText(text As String) As String
            If text Is Nothing Then
                Return String.Empty
            End If

            Return text.Trim()
        End Function

        Private Async Function HandleServerRequestAsync(request As RpcServerRequest) As Task
            Dim approvalInfo As PendingApprovalInfo = Nothing
            If _approvalService.TryCreateApproval(request, approvalInfo) Then
                QueueApproval(approvalInfo)
                Return
            End If

            Select Case request.MethodName
                Case "item/tool/requestUserInput"
                    Await HandleToolRequestUserInputAsync(request)

                Case "item/tool/call"
                    Await HandleUnsupportedToolCallAsync(request)

                Case "account/chatgptAuthTokens/refresh"
                    Await HandleChatgptTokenRefreshAsync(request)

                Case Else
                    Await CurrentClient().SendErrorAsync(request.Id, -32601, $"Unsupported server request method: {request.MethodName}")
            End Select
        End Function

        Private Sub QueueApproval(approvalInfo As PendingApprovalInfo)
            If approvalInfo Is Nothing Then
                Return
            End If

            _approvalQueue.Enqueue(approvalInfo)
            ShowNextApprovalIfNeeded()
            AppendSystemMessage($"Approval queued: {approvalInfo.MethodName}")
            ShowStatus($"Approval queued: {approvalInfo.MethodName}", displayToast:=True)
        End Sub

        Private Sub ShowNextApprovalIfNeeded()
            If _activeApproval IsNot Nothing Then
                Return
            End If

            If _approvalQueue.Count = 0 Then
                TxtApproval.Text = "No pending approvals."
                RefreshControlStates()
                Return
            End If

            _activeApproval = _approvalQueue.Dequeue()
            TxtApproval.Text = _activeApproval.Summary
            RefreshControlStates()
        End Sub

        Private Async Function ResolveApprovalAsync(action As String) As Task
            If _activeApproval Is Nothing Then
                Return
            End If

            Dim decision = _approvalService.ResolveDecision(_activeApproval, action)
            If String.IsNullOrWhiteSpace(decision) Then
                Throw New InvalidOperationException("No decision mapping is available for this approval type.")
            End If

            Dim resultNode As New JsonObject()
            resultNode("decision") = decision

            Await CurrentClient().SendResultAsync(_activeApproval.RequestId, resultNode)
            AppendSystemMessage($"Approval sent: {decision}")
            ShowStatus($"Approval sent: {decision}", displayToast:=True)

            _activeApproval = Nothing
            ShowNextApprovalIfNeeded()
        End Function

        Private Async Function HandleToolRequestUserInputAsync(request As RpcServerRequest) As Task
            Dim paramsObject = AsObject(request.ParamsNode)
            Dim questions = GetPropertyArray(paramsObject, "questions")
            If questions Is Nothing OrElse questions.Count = 0 Then
                Await CurrentClient().SendErrorAsync(request.Id, -32602, "No questions were provided in item/tool/requestUserInput.")
                Return
            End If

            Dim answersRoot As New JsonObject()

            For Each questionNode In questions
                Dim questionObject = AsObject(questionNode)
                If questionObject Is Nothing Then
                    Continue For
                End If

                Dim questionId = GetPropertyString(questionObject, "id")
                If String.IsNullOrWhiteSpace(questionId) Then
                    Continue For
                End If

                Dim header = GetPropertyString(questionObject, "header", "Input required")
                Dim prompt = GetPropertyString(questionObject, "question", questionId)
                Dim isSecret = GetPropertyBoolean(questionObject, "isSecret", False)

                Dim options As New List(Of String)()
                Dim optionsArray = GetPropertyArray(questionObject, "options")
                If optionsArray IsNot Nothing Then
                    For Each optionNode In optionsArray
                        Dim optionObject = AsObject(optionNode)
                        If optionObject Is Nothing Then
                            Continue For
                        End If

                        Dim label = GetPropertyString(optionObject, "label")
                        Dim description = GetPropertyString(optionObject, "description")
                        If String.IsNullOrWhiteSpace(label) Then
                            Continue For
                        End If

                        If String.IsNullOrWhiteSpace(description) Then
                            options.Add(label)
                        Else
                            options.Add($"{label} - {description}")
                        End If
                    Next
                End If

                Dim answer As String = Nothing
                Dim promptDialog As New QuestionPromptWindow(header, prompt, options, isSecret) With {
                    .Owner = Me
                }
                Dim dialogResult = promptDialog.ShowDialog()
                If Not dialogResult.HasValue OrElse Not dialogResult.Value Then
                    Await CurrentClient().SendErrorAsync(request.Id, -32800, "User canceled request_user_input.")
                    Return
                End If

                answer = promptDialog.Answer

                If String.IsNullOrWhiteSpace(answer) Then
                    Await CurrentClient().SendErrorAsync(request.Id, -32602, $"No answer provided for question '{questionId}'.")
                    Return
                End If

                Dim answerObject As New JsonObject()
                Dim answerList As New JsonArray()
                answerList.Add(answer)
                answerObject("answers") = answerList
                answersRoot(questionId) = answerObject
            Next

            Dim response As New JsonObject()
            response("answers") = answersRoot

            Await CurrentClient().SendResultAsync(request.Id, response)
            AppendSystemMessage("Submitted request_user_input answers.")
            ShowStatus("Submitted request_user_input answers.", displayToast:=True)
        End Function

        Private Async Function HandleUnsupportedToolCallAsync(request As RpcServerRequest) As Task
            Dim response As New JsonObject()
            response("success") = False

            Dim contentItems As New JsonArray()
            Dim textItem As New JsonObject()
            textItem("type") = "inputText"
            textItem("text") = "This host does not expose dynamic tool callbacks yet."
            contentItems.Add(textItem)

            response("contentItems") = contentItems
            Await CurrentClient().SendResultAsync(request.Id, response)
            AppendSystemMessage("Declined dynamic tool call (unsupported in this host).")
            ShowStatus("Declined dynamic tool call.", isError:=True)
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
            _lastActivityUtc = DateTimeOffset.UtcNow
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

            _reconnectInProgress = False
            _reconnectAttempt = 0
            _nextReconnectAttemptUtc = Nothing
            UpdateReconnectCountdownUi()
        End Sub

        Private Async Function ReconnectNowAsync() As Task
            If _client IsNot Nothing AndAlso _client.IsRunning Then
                ShowStatus("Already connected.")
                Return
            End If

            CancelReconnect()
            BeginReconnect("manual reconnect requested", force:=True)
            Await Task.CompletedTask
        End Function

        Private Sub BeginReconnect(reason As String, Optional force As Boolean = False)
            If Not force AndAlso Not _connectionExpected Then
                Return
            End If

            If Not force AndAlso Not IsChecked(ChkAutoReconnect) Then
                _connectionExpected = False
                ShowStatus("Disconnected. Auto-reconnect is disabled.", isError:=True)
                UpdateReconnectCountdownUi()
                Return
            End If

            If _reconnectInProgress Then
                If force Then
                    CancelReconnect()
                Else
                    Return
                End If
            End If

            If _reconnectInProgress Then
                Return
            End If

            _connectionExpected = True

            Dim reconnectCts As New CancellationTokenSource()
            _reconnectCts = reconnectCts
            _reconnectInProgress = True
            _reconnectAttempt = 0
            _nextReconnectAttemptUtc = Nothing
            ShowStatus($"Auto-reconnect scheduled: {reason}", isError:=True, displayToast:=True)
            AppendSystemMessage($"Auto-reconnect scheduled: {reason}")
            UpdateReconnectCountdownUi()

            FireAndForget(ReconnectLoopAsync(reason, reconnectCts))
        End Sub

        Private Async Function ReconnectLoopAsync(reason As String,
                                                  reconnectCts As CancellationTokenSource) As Task
            Dim token = reconnectCts.Token
            Dim delays As Integer() = {2, 5, 10, 20, 30}

            Try
                For attempt = 1 To delays.Length + 1
                    token.ThrowIfCancellationRequested()

                    Dim currentAttempt = attempt
                    _reconnectAttempt = currentAttempt
                    _nextReconnectAttemptUtc = Nothing
                    RunOnUi(
                        Sub()
                            ShowStatus($"Reconnect attempt {currentAttempt}/{delays.Length + 1}: {reason}", isError:=True)
                            UpdateReconnectCountdownUi()
                        End Sub)

                    Dim delayBeforeRetry As TimeSpan? = Nothing
                    Try
                        Await RunOnUiAsync(Function() ConnectCoreAsync(isReconnect:=True, cancellationToken:=token))

                        RunOnUi(
                            Sub()
                                ShowStatus("Reconnected successfully.", displayToast:=True)
                                AppendSystemMessage("Reconnected successfully.")
                            End Sub)

                        Return
                    Catch ex As OperationCanceledException
                        Throw
                    Catch ex As Exception
                        RunOnUi(
                            Sub()
                                AppendSystemMessage($"Reconnect attempt {currentAttempt} failed: {ex.Message}")
                                ShowStatus($"Reconnect attempt {currentAttempt} failed.", isError:=True)
                            End Sub)

                        If currentAttempt <= delays.Length Then
                            delayBeforeRetry = TimeSpan.FromSeconds(delays(currentAttempt - 1))
                        End If
                    End Try

                    If delayBeforeRetry.HasValue Then
                        _nextReconnectAttemptUtc = DateTimeOffset.UtcNow.Add(delayBeforeRetry.Value)
                        RunOnUi(Sub() UpdateReconnectCountdownUi())
                        Await Task.Delay(delayBeforeRetry.Value, token)
                    End If
                Next

                RunOnUi(
                    Sub()
                        _connectionExpected = False
                        _nextReconnectAttemptUtc = Nothing
                        UpdateReconnectCountdownUi()
                        ShowStatus("Auto-reconnect failed. Please reconnect manually.",
                                   isError:=True,
                                   displayToast:=True)
                    End Sub)
            Catch ex As OperationCanceledException
                RunOnUi(
                    Sub()
                        _nextReconnectAttemptUtc = Nothing
                        UpdateReconnectCountdownUi()
                        ShowStatus("Auto-reconnect canceled.")
                    End Sub)
            Finally
                RunOnUi(
                    Sub()
                        If ReferenceEquals(_reconnectCts, reconnectCts) Then
                            _reconnectCts = Nothing
                        End If

                        reconnectCts.Dispose()
                        _reconnectInProgress = False
                        _reconnectAttempt = 0
                        _nextReconnectAttemptUtc = Nothing
                        UpdateReconnectCountdownUi()
                    End Sub)
            End Try
        End Function

        Private Sub OnWatchdogTimerTick(sender As Object, e As EventArgs)
            If Not _connectionExpected Then
                Return
            End If

            Dim connected = _client IsNot Nothing AndAlso _client.IsRunning
            If Not connected Then
                BeginReconnect("watchdog detected a disconnected client")
                Return
            End If

            Dim inactiveFor = DateTimeOffset.UtcNow - _lastActivityUtc
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
            If _client IsNot Nothing AndAlso _client.IsRunning Then
                LblReconnectCountdown.Text = "Reconnect: connected."
                Return
            End If

            If _reconnectInProgress Then
                If _nextReconnectAttemptUtc.HasValue Then
                    Dim remaining = _nextReconnectAttemptUtc.Value - DateTimeOffset.UtcNow
                    Dim secondsLeft = Math.Max(0, CInt(Math.Ceiling(remaining.TotalSeconds)))
                    LblReconnectCountdown.Text = $"Reconnect: next attempt in {secondsLeft}s."
                    Return
                End If

                Dim attemptText = If(_reconnectAttempt > 0, _reconnectAttempt.ToString(), "?")
                LblReconnectCountdown.Text = $"Reconnect: attempt {attemptText} running..."
                Return
            End If

            If _connectionExpected AndAlso IsChecked(ChkAutoReconnect) Then
                LblReconnectCountdown.Text = "Reconnect: standing by."
                Return
            End If

            LblReconnectCountdown.Text = "Reconnect: not scheduled."
        End Sub

        Private Async Function ExportDiagnosticsAsync() As Task
            SaveSettings()

            Dim diagnosticsRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                                               "CodexNativeAgentDiagnostics")
            Directory.CreateDirectory(diagnosticsRoot)

            Dim stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss")
            Dim outputFolder = Path.Combine(diagnosticsRoot, $"diag-{stamp}")
            Directory.CreateDirectory(outputFolder)

            File.WriteAllText(Path.Combine(outputFolder, "transcript.log"), TxtTranscript.Text)
            File.WriteAllText(Path.Combine(outputFolder, "protocol.log"), TxtProtocol.Text)
            File.WriteAllText(Path.Combine(outputFolder, "approval.txt"), TxtApproval.Text)
            File.WriteAllText(Path.Combine(outputFolder, "rate-limits.txt"), TxtRateLimits.Text)

            Dim threadBuilder As New StringBuilder()
            For Each entry In _threadEntries
                threadBuilder.AppendLine($"{entry.Id}{ControlChars.Tab}{entry.LastActiveAt}{ControlChars.Tab}{entry.Cwd}{ControlChars.Tab}{entry.Preview}")
            Next
            File.WriteAllText(Path.Combine(outputFolder, "threads.tsv"), threadBuilder.ToString())

            Dim settingsObject As New JsonObject()
            settingsObject("codexPath") = TxtCodexPath.Text.Trim()
            settingsObject("serverArgs") = TxtServerArgs.Text.Trim()
            settingsObject("workingDir") = TxtWorkingDir.Text.Trim()
            settingsObject("windowsCodexHome") = TxtWindowsCodexHome.Text.Trim()
            settingsObject("rememberApiKey") = IsChecked(ChkRememberApiKey)
            settingsObject("autoLoginApiKey") = IsChecked(ChkAutoLoginApiKey)
            settingsObject("autoReconnect") = IsChecked(ChkAutoReconnect)
            settingsObject("theme") = _currentTheme
            settingsObject("density") = _currentDensity
            settingsObject("apiKeyMasked") = MaskSecret(TxtApiKey.Text.Trim(), 4)
            settingsObject("externalAccessTokenMasked") = MaskSecret(TxtExternalAccessToken.Text.Trim(), 6)
            settingsObject("externalAccountIdMasked") = MaskSecret(TxtExternalAccountId.Text.Trim(), 4)

            Dim connectionObject As New JsonObject()
            connectionObject("isConnected") = (_client IsNot Nothing AndAlso _client.IsRunning)
            connectionObject("expectedConnection") = _connectionExpected
            connectionObject("reconnectInProgress") = _reconnectInProgress
            connectionObject("reconnectAttempt") = _reconnectAttempt
            connectionObject("currentThreadId") = _currentThreadId
            connectionObject("currentTurnId") = _currentTurnId
            connectionObject("lastActivityUtc") = _lastActivityUtc.ToString("O")
            connectionObject("processId") = If(_client Is Nothing, 0, _client.ProcessId)

            Dim uiObject As New JsonObject()
            uiObject("threadSearch") = TxtThreadSearch.Text
            uiObject("threadSort") = SelectedComboValue(CmbThreadSort)
            uiObject("model") = SelectedModelId()
            uiObject("approvalPolicy") = SelectedComboValue(CmbApprovalPolicy)
            uiObject("sandbox") = SelectedComboValue(CmbSandbox)
            uiObject("reasoningEffort") = SelectedComboValue(CmbReasoningEffort)
            uiObject("theme") = _currentTheme
            uiObject("density") = _currentDensity

            Dim snapshot As New JsonObject()
            snapshot("generatedAtLocal") = DateTime.Now.ToString("O")
            snapshot("generatedAtUtc") = DateTimeOffset.UtcNow.ToString("O")
            snapshot("statusText") = LblStatus.Text
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
            textBox.CaretIndex = textBox.Text.Length
            textBox.ScrollToEnd()
        End Sub

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
