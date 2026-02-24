Imports System.IO
Imports System.Windows.Controls

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private Sub ToggleTheme()
            Dim nextTheme = AppAppearanceManager.ToggleTheme(_currentTheme)
            ApplyAppearance(nextTheme, _currentDensity, persist:=True)
            ShowStatus($"Theme switched to {AppAppearanceManager.DisplayTheme(_currentTheme)}.", displayToast:=True)
        End Sub

        Private Sub OnDensitySelectionChanged()
            If _suppressAppearanceUiChange Then
                Return
            End If

            Dim selectedDensity = AppAppearanceManager.NormalizeDensity(DensityValueFromIndex(_viewModel.SettingsPanel.DensityIndex))
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
                _viewModel.SettingsPanel.DensityIndex = If(compact, 1, 0)
                _viewModel.SettingsPanel.ThemeStateText = $"Current: {AppAppearanceManager.DisplayTheme(_currentTheme)}"
                _viewModel.SettingsPanel.ThemeToggleButtonText = AppAppearanceManager.ThemeButtonLabel(_currentTheme)
            Finally
                _suppressAppearanceUiChange = False
            End Try
        End Sub

        Private Function LoadSettings() As AppSettings
            If _settingsStore Is Nothing Then
                Return New AppSettings()
            End If

            Return _settingsStore.Load()
        End Function

        Private Sub SaveSettings()
            CaptureSettingsFromControls()

            If Not _settings.RememberApiKey Then
                _settings.EncryptedApiKey = String.Empty
            End If

            If _settingsStore IsNot Nothing Then
                _settingsStore.Save(_settings)
            End If
        End Sub

        Private Sub CaptureSettingsFromControls()
            _settings.CodexPath = _viewModel.SettingsPanel.CodexPath.Trim()
            _settings.ServerArgs = _viewModel.SettingsPanel.ServerArgs.Trim()
            _settings.WorkingDir = _viewModel.SettingsPanel.WorkingDir.Trim()
            _settings.WindowsCodexHome = _viewModel.SettingsPanel.WindowsCodexHome.Trim()
            _settings.RememberApiKey = _viewModel.SettingsPanel.RememberApiKey
            _settings.AutoLoginApiKey = _viewModel.SettingsPanel.AutoLoginApiKey
            _settings.AutoReconnect = _viewModel.SettingsPanel.AutoReconnect
            _settings.DisableWorkspaceHintOverlay = _viewModel.SettingsPanel.DisableWorkspaceHintOverlay
            _settings.DisableConnectionInitializedToast = _viewModel.SettingsPanel.DisableConnectionInitializedToast
            _settings.DisableThreadsPanelHints = _viewModel.SettingsPanel.DisableThreadsPanelHints
            _settings.ShowEventDotsInTranscript = _viewModel.SettingsPanel.ShowEventDotsInTranscript
            _settings.ShowSystemDotsInTranscript = _viewModel.SettingsPanel.ShowSystemDotsInTranscript
            _settings.FilterThreadsByWorkingDir = _viewModel.ThreadsPanel.FilterByWorkingDir
            _settings.ThemeMode = _currentTheme
            _settings.DensityMode = _currentDensity
            _settings.TurnComposerPickersCollapsed = _turnComposerPickersCollapsed
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

            _viewModel.SettingsPanel.CodexPath = _settings.CodexPath
            _viewModel.SettingsPanel.ServerArgs = _settings.ServerArgs
            _viewModel.SettingsPanel.WorkingDir = _settings.WorkingDir
            _viewModel.SettingsPanel.WindowsCodexHome = _settings.WindowsCodexHome
            _viewModel.SettingsPanel.RememberApiKey = _settings.RememberApiKey
            _viewModel.SettingsPanel.AutoLoginApiKey = _settings.AutoLoginApiKey
            _viewModel.SettingsPanel.AutoReconnect = _settings.AutoReconnect
            _viewModel.SettingsPanel.DisableWorkspaceHintOverlay = _settings.DisableWorkspaceHintOverlay
            _viewModel.SettingsPanel.DisableConnectionInitializedToast = _settings.DisableConnectionInitializedToast
            _viewModel.SettingsPanel.DisableThreadsPanelHints = _settings.DisableThreadsPanelHints
            _viewModel.SettingsPanel.ShowEventDotsInTranscript = _settings.ShowEventDotsInTranscript
            _viewModel.SettingsPanel.ShowSystemDotsInTranscript = _settings.ShowSystemDotsInTranscript
            _viewModel.ThreadsPanel.FilterByWorkingDir = _settings.FilterThreadsByWorkingDir
            _turnComposerPickersCollapsed = _settings.TurnComposerPickersCollapsed
            ApplyAppearance(_settings.ThemeMode, _settings.DensityMode, persist:=False)
            ApplyTranscriptTimelineDotVisibilitySettings()

            If _viewModel.SettingsPanel.RememberApiKey Then
                Dim decryptedApiKey = ReadPersistedApiKey()
                If Not String.IsNullOrWhiteSpace(decryptedApiKey) Then
                    _viewModel.SettingsPanel.ApiKey = decryptedApiKey
                End If
            End If

            UpdateRuntimeFieldState()
            _viewModel.ApprovalPanel.ResetLifecycleState()
            _viewModel.SettingsPanel.RateLimitsText = "No rate-limit data loaded yet."
            SyncThreadToolbarMenus()
            SyncNewThreadTargetChip()
            ApplyTurnComposerPickersCollapsedState(animated:=False, persist:=False)
        End Sub

        Private Sub ApplyTranscriptTimelineDotVisibilitySettings()
            If _viewModel Is Nothing OrElse _viewModel.TranscriptPanel Is Nothing OrElse _viewModel.SettingsPanel Is Nothing Then
                Return
            End If

            _viewModel.TranscriptPanel.ShowEventDotsInTranscript = _viewModel.SettingsPanel.ShowEventDotsInTranscript
            _viewModel.TranscriptPanel.ShowSystemDotsInTranscript = _viewModel.SettingsPanel.ShowSystemDotsInTranscript
            UpdateWorkspaceEmptyStateVisibility()
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
    End Class
End Namespace
