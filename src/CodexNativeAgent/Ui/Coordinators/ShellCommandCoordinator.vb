Imports System.Threading.Tasks
Imports System.Windows.Input
Imports CodexNativeAgent.Ui.Mvvm
Imports CodexNativeAgent.Ui.ViewModels

Namespace CodexNativeAgent.Ui.Coordinators
    Public NotInheritable Class ShellCommandCoordinator
        Private ReadOnly _viewModel As MainWindowViewModel
        Private ReadOnly _runUiActionAsync As Func(Of Func(Of Task), Task)
        Private ReadOnly _fireAndForget As Action(Of Task)

        Public Sub New(viewModel As MainWindowViewModel,
                       runUiActionAsync As Func(Of Func(Of Task), Task),
                       fireAndForget As Action(Of Task))
            If viewModel Is Nothing Then Throw New ArgumentNullException(NameOf(viewModel))
            If runUiActionAsync Is Nothing Then Throw New ArgumentNullException(NameOf(runUiActionAsync))
            If fireAndForget Is Nothing Then Throw New ArgumentNullException(NameOf(fireAndForget))

            _viewModel = viewModel
            _runUiActionAsync = runUiActionAsync
            _fireAndForget = fireAndForget
        End Sub

        Public Sub BindCommands(startTurnAsync As Func(Of Task),
                                refreshThreadsAsync As Func(Of Task),
                                refreshModelsAsync As Func(Of Task),
                                startThreadAsync As Func(Of Task),
                                focusThreadSearch As Action,
                                openSettings As Action,
                                openSortMenu As Action,
                                openFilterMenu As Action,
                                setSortFromMenu As Action(Of Integer),
                                applyFilterMenuToggle As Action,
                                selectThreadContext As Action,
                                refreshThreadContextAsync As Func(Of Task),
                                forkThreadContextAsync As Func(Of Task),
                                archiveThreadContextAsync As Func(Of Task),
                                unarchiveThreadContextAsync As Func(Of Task),
                                newThreadFromGroupAsync As Func(Of Task),
                                openVsCodeFromGroupAsync As Func(Of Task),
                                openTerminalFromGroupAsync As Func(Of Task),
                                toggleTheme As Action,
                                exportDiagnosticsAsync As Func(Of Task))
            If startTurnAsync Is Nothing Then Throw New ArgumentNullException(NameOf(startTurnAsync))
            If refreshThreadsAsync Is Nothing Then Throw New ArgumentNullException(NameOf(refreshThreadsAsync))
            If refreshModelsAsync Is Nothing Then Throw New ArgumentNullException(NameOf(refreshModelsAsync))
            If startThreadAsync Is Nothing Then Throw New ArgumentNullException(NameOf(startThreadAsync))
            If focusThreadSearch Is Nothing Then Throw New ArgumentNullException(NameOf(focusThreadSearch))
            If openSettings Is Nothing Then Throw New ArgumentNullException(NameOf(openSettings))
            If openSortMenu Is Nothing Then Throw New ArgumentNullException(NameOf(openSortMenu))
            If openFilterMenu Is Nothing Then Throw New ArgumentNullException(NameOf(openFilterMenu))
            If setSortFromMenu Is Nothing Then Throw New ArgumentNullException(NameOf(setSortFromMenu))
            If applyFilterMenuToggle Is Nothing Then Throw New ArgumentNullException(NameOf(applyFilterMenuToggle))
            If selectThreadContext Is Nothing Then Throw New ArgumentNullException(NameOf(selectThreadContext))
            If refreshThreadContextAsync Is Nothing Then Throw New ArgumentNullException(NameOf(refreshThreadContextAsync))
            If forkThreadContextAsync Is Nothing Then Throw New ArgumentNullException(NameOf(forkThreadContextAsync))
            If archiveThreadContextAsync Is Nothing Then Throw New ArgumentNullException(NameOf(archiveThreadContextAsync))
            If unarchiveThreadContextAsync Is Nothing Then Throw New ArgumentNullException(NameOf(unarchiveThreadContextAsync))
            If newThreadFromGroupAsync Is Nothing Then Throw New ArgumentNullException(NameOf(newThreadFromGroupAsync))
            If openVsCodeFromGroupAsync Is Nothing Then Throw New ArgumentNullException(NameOf(openVsCodeFromGroupAsync))
            If openTerminalFromGroupAsync Is Nothing Then Throw New ArgumentNullException(NameOf(openTerminalFromGroupAsync))
            If toggleTheme Is Nothing Then Throw New ArgumentNullException(NameOf(toggleTheme))
            If exportDiagnosticsAsync Is Nothing Then Throw New ArgumentNullException(NameOf(exportDiagnosticsAsync))

            BindShellCommands(startTurnAsync,
                              refreshThreadsAsync,
                              refreshModelsAsync,
                              startThreadAsync,
                              focusThreadSearch,
                              openSettings)

            BindThreadPanelCommands(openSortMenu,
                                    openFilterMenu,
                                    setSortFromMenu,
                                    applyFilterMenuToggle,
                                    selectThreadContext,
                                    refreshThreadContextAsync,
                                    forkThreadContextAsync,
                                    archiveThreadContextAsync,
                                    unarchiveThreadContextAsync,
                                    newThreadFromGroupAsync,
                                    openVsCodeFromGroupAsync,
                                    openTerminalFromGroupAsync)

            _viewModel.SettingsPanel.ToggleThemeCommand = New RelayCommand(
                Sub()
                    toggleTheme()
                End Sub)

            _viewModel.SettingsPanel.ExportDiagnosticsCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(exportDiagnosticsAsync)
                End Function)
        End Sub

        Private Sub BindShellCommands(startTurnAsync As Func(Of Task),
                                      refreshThreadsAsync As Func(Of Task),
                                      refreshModelsAsync As Func(Of Task),
                                      startThreadAsync As Func(Of Task),
                                      focusThreadSearch As Action,
                                      openSettings As Action)
            _viewModel.ShellSendCommand = New RelayCommand(
                Sub()
                    _fireAndForget(_runUiActionAsync(startTurnAsync))
                End Sub,
                Function() _viewModel.TurnComposer.CanStartTurn)

            _viewModel.ShellRefreshThreadsCommand = New RelayCommand(
                Sub()
                    _fireAndForget(_runUiActionAsync(refreshThreadsAsync))
                End Sub,
                Function() _viewModel.ThreadsPanel.CanRunFullRefresh)

            _viewModel.ShellRefreshModelsCommand = New RelayCommand(
                Sub()
                    _fireAndForget(_runUiActionAsync(refreshModelsAsync))
                End Sub,
                Function() _viewModel.SessionState.IsAuthenticated)

            _viewModel.ShellNewThreadCommand = New RelayCommand(
                Sub()
                    _fireAndForget(_runUiActionAsync(startThreadAsync))
                End Sub,
                Function() _viewModel.SessionState.IsAuthenticated AndAlso Not _viewModel.ThreadsPanel.IsThreadContentLoading)

            _viewModel.ShellFocusThreadSearchCommand = New RelayCommand(
                Sub()
                    focusThreadSearch()
                End Sub)

            _viewModel.ShellOpenSettingsCommand = New RelayCommand(
                Sub()
                    openSettings()
                End Sub)
        End Sub

        Private Sub BindThreadPanelCommands(openSortMenu As Action,
                                            openFilterMenu As Action,
                                            setSortFromMenu As Action(Of Integer),
                                            applyFilterMenuToggle As Action,
                                            selectThreadContext As Action,
                                            refreshThreadContextAsync As Func(Of Task),
                                            forkThreadContextAsync As Func(Of Task),
                                            archiveThreadContextAsync As Func(Of Task),
                                            unarchiveThreadContextAsync As Func(Of Task),
                                            newThreadFromGroupAsync As Func(Of Task),
                                            openVsCodeFromGroupAsync As Func(Of Task),
                                            openTerminalFromGroupAsync As Func(Of Task))
            _viewModel.ThreadsPanel.OpenSortMenuCommand = New RelayCommand(
                Sub()
                    openSortMenu()
                End Sub)

            _viewModel.ThreadsPanel.OpenFilterMenuCommand = New RelayCommand(
                Sub()
                    openFilterMenu()
                End Sub)

            _viewModel.ThreadsPanel.SetSortFromMenuCommand = New RelayCommand(
                Sub(parameter)
                    Dim parsed As Integer
                    If Not Integer.TryParse(If(parameter, String.Empty).ToString(), parsed) Then
                        Return
                    End If

                    setSortFromMenu(parsed)
                End Sub)

            _viewModel.ThreadsPanel.ApplyFilterMenuToggleCommand = New RelayCommand(
                Sub()
                    applyFilterMenuToggle()
                End Sub)

            _viewModel.ThreadsPanel.SelectThreadContextCommand = New RelayCommand(
                Sub()
                    selectThreadContext()
                End Sub)

            _viewModel.ThreadsPanel.RefreshThreadContextCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(refreshThreadContextAsync)
                End Function)

            _viewModel.ThreadsPanel.ForkThreadContextCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(forkThreadContextAsync)
                End Function)

            _viewModel.ThreadsPanel.ArchiveThreadContextCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(archiveThreadContextAsync)
                End Function)

            _viewModel.ThreadsPanel.UnarchiveThreadContextCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(unarchiveThreadContextAsync)
                End Function)

            _viewModel.ThreadsPanel.NewThreadFromGroupContextCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(newThreadFromGroupAsync)
                End Function)

            _viewModel.ThreadsPanel.OpenVsCodeFromGroupContextCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(openVsCodeFromGroupAsync)
                End Function)

            _viewModel.ThreadsPanel.OpenTerminalFromGroupContextCommand = New AsyncRelayCommand(
                Function()
                    Return _runUiActionAsync(openTerminalFromGroupAsync)
                End Function)
        End Sub

    End Class
End Namespace
