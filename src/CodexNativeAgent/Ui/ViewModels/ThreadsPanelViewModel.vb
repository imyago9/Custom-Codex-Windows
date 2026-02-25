Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Input
Imports CodexNativeAgent.Ui.Mvvm

Namespace CodexNativeAgent.Ui.ViewModels
    Public NotInheritable Class ThreadsPanelViewModel
        Inherits ViewModelBase

        Private _searchText As String = String.Empty
        Private _sortIndex As Integer
        Private _showArchived As Boolean
        Private _filterByWorkingDir As Boolean
        Private _isSearchEnabled As Boolean
        Private _isSortMenuEnabled As Boolean
        Private _isFilterMenuEnabled As Boolean
        Private _isThreadListEnabled As Boolean
        Private _isLoading As Boolean
        Private _isThreadContentLoading As Boolean
        Private _canRunFullRefresh As Boolean
        Private _canAutoLoadSelection As Boolean
        Private _canRunThreadContextActions As Boolean
        Private _selectionLoadVersion As Integer
        Private _selectionLoadThreadId As String = String.Empty
        Private _refreshErrorText As String = String.Empty
        Private _lastRefreshStartedUtc As DateTimeOffset?
        Private _lastRefreshCompletedUtc As DateTimeOffset?
        Private _lastRefreshThreadCount As Integer = -1
        Private _lastSelectionLoadStartedUtc As DateTimeOffset?
        Private _lastSelectionLoadCompletedUtc As DateTimeOffset?
        Private _lastSelectionLoadErrorText As String = String.Empty
        Private _stateText As String = "No threads loaded yet."
        Private _selectedListItem As Object
        Private ReadOnly _items As New ObservableCollection(Of Object)()
        Private _threadMenuGroupNewThreadVisibility As Visibility = Visibility.Collapsed
        Private _threadMenuGroupOpenVsCodeVisibility As Visibility = Visibility.Collapsed
        Private _threadMenuGroupOpenTerminalVisibility As Visibility = Visibility.Collapsed
        Private _threadMenuGroupTopSeparatorVisibility As Visibility = Visibility.Collapsed
        Private _threadMenuSelectVisibility As Visibility = Visibility.Visible
        Private _threadMenuSelectRefreshSeparatorVisibility As Visibility = Visibility.Visible
        Private _threadMenuRefreshVisibility As Visibility = Visibility.Visible
        Private _threadMenuRefreshActionsSeparatorVisibility As Visibility = Visibility.Visible
        Private _threadMenuForkVisibility As Visibility = Visibility.Visible
        Private _threadMenuArchiveVisibility As Visibility = Visibility.Visible
        Private _threadMenuUnarchiveVisibility As Visibility = Visibility.Visible
        Private _threadMenuGroupNewThreadEnabled As Boolean
        Private _threadMenuGroupOpenVsCodeEnabled As Boolean
        Private _threadMenuGroupOpenTerminalEnabled As Boolean
        Private _threadMenuSelectEnabled As Boolean = True
        Private _threadMenuRefreshEnabled As Boolean
        Private _threadMenuForkEnabled As Boolean
        Private _threadMenuArchiveEnabled As Boolean
        Private _threadMenuUnarchiveEnabled As Boolean
        Private _openSortMenuCommand As ICommand
        Private _openFilterMenuCommand As ICommand
        Private _setSortFromMenuCommand As ICommand
        Private _applyFilterMenuToggleCommand As ICommand
        Private _selectThreadContextCommand As ICommand
        Private _refreshThreadContextCommand As ICommand
        Private _forkThreadContextCommand As ICommand
        Private _archiveThreadContextCommand As ICommand
        Private _unarchiveThreadContextCommand As ICommand
        Private _newThreadFromGroupContextCommand As ICommand
        Private _openVsCodeFromGroupContextCommand As ICommand
        Private _openTerminalFromGroupContextCommand As ICommand

        Public Property SearchText As String
            Get
                Return _searchText
            End Get
            Set(value As String)
                SetProperty(_searchText, If(value, String.Empty))
            End Set
        End Property

        Public Property SortIndex As Integer
            Get
                Return _sortIndex
            End Get
            Set(value As Integer)
                SetProperty(_sortIndex, Math.Max(0, value))
            End Set
        End Property

        Public Property ShowArchived As Boolean
            Get
                Return _showArchived
            End Get
            Set(value As Boolean)
                SetProperty(_showArchived, value)
            End Set
        End Property

        Public Property FilterByWorkingDir As Boolean
            Get
                Return _filterByWorkingDir
            End Get
            Set(value As Boolean)
                SetProperty(_filterByWorkingDir, value)
            End Set
        End Property

        Public Property IsSearchEnabled As Boolean
            Get
                Return _isSearchEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isSearchEnabled, value)
            End Set
        End Property

        Public Property IsSortMenuEnabled As Boolean
            Get
                Return _isSortMenuEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isSortMenuEnabled, value)
            End Set
        End Property

        Public Property IsFilterMenuEnabled As Boolean
            Get
                Return _isFilterMenuEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isFilterMenuEnabled, value)
            End Set
        End Property

        Public Property IsThreadListEnabled As Boolean
            Get
                Return _isThreadListEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isThreadListEnabled, value)
            End Set
        End Property

        Public Property IsLoading As Boolean
            Get
                Return _isLoading
            End Get
            Set(value As Boolean)
                SetProperty(_isLoading, value)
            End Set
        End Property

        Public Property IsThreadContentLoading As Boolean
            Get
                Return _isThreadContentLoading
            End Get
            Set(value As Boolean)
                SetProperty(_isThreadContentLoading, value)
            End Set
        End Property

        Public Property CanRunFullRefresh As Boolean
            Get
                Return _canRunFullRefresh
            End Get
            Set(value As Boolean)
                SetProperty(_canRunFullRefresh, value)
            End Set
        End Property

        Public Property CanAutoLoadSelection As Boolean
            Get
                Return _canAutoLoadSelection
            End Get
            Set(value As Boolean)
                SetProperty(_canAutoLoadSelection, value)
            End Set
        End Property

        Public Property CanRunThreadContextActions As Boolean
            Get
                Return _canRunThreadContextActions
            End Get
            Set(value As Boolean)
                SetProperty(_canRunThreadContextActions, value)
            End Set
        End Property

        Public ReadOnly Property SelectionLoadVersion As Integer
            Get
                Return _selectionLoadVersion
            End Get
        End Property

        Public ReadOnly Property SelectionLoadThreadId As String
            Get
                Return _selectionLoadThreadId
            End Get
        End Property

        Public ReadOnly Property HasActiveSelectionLoad As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(_selectionLoadThreadId)
            End Get
        End Property

        Public Property RefreshErrorText As String
            Get
                Return _refreshErrorText
            End Get
            Set(value As String)
                SetProperty(_refreshErrorText, If(value, String.Empty))
            End Set
        End Property

        Public Property LastRefreshStartedUtc As DateTimeOffset?
            Get
                Return _lastRefreshStartedUtc
            End Get
            Set(value As DateTimeOffset?)
                SetProperty(_lastRefreshStartedUtc, value)
            End Set
        End Property

        Public Property LastRefreshCompletedUtc As DateTimeOffset?
            Get
                Return _lastRefreshCompletedUtc
            End Get
            Set(value As DateTimeOffset?)
                SetProperty(_lastRefreshCompletedUtc, value)
            End Set
        End Property

        Public Property LastRefreshThreadCount As Integer
            Get
                Return _lastRefreshThreadCount
            End Get
            Set(value As Integer)
                SetProperty(_lastRefreshThreadCount, value)
            End Set
        End Property

        Public Property LastSelectionLoadStartedUtc As DateTimeOffset?
            Get
                Return _lastSelectionLoadStartedUtc
            End Get
            Set(value As DateTimeOffset?)
                SetProperty(_lastSelectionLoadStartedUtc, value)
            End Set
        End Property

        Public Property LastSelectionLoadCompletedUtc As DateTimeOffset?
            Get
                Return _lastSelectionLoadCompletedUtc
            End Get
            Set(value As DateTimeOffset?)
                SetProperty(_lastSelectionLoadCompletedUtc, value)
            End Set
        End Property

        Public Property LastSelectionLoadErrorText As String
            Get
                Return _lastSelectionLoadErrorText
            End Get
            Set(value As String)
                SetProperty(_lastSelectionLoadErrorText, If(value, String.Empty))
            End Set
        End Property

        Public Property StateText As String
            Get
                Return _stateText
            End Get
            Set(value As String)
                SetProperty(_stateText, If(value, String.Empty))
            End Set
        End Property

        Public ReadOnly Property Items As ObservableCollection(Of Object)
            Get
                Return _items
            End Get
        End Property

        Public Property SelectedListItem As Object
            Get
                Return _selectedListItem
            End Get
            Set(value As Object)
                SetProperty(_selectedListItem, value)
            End Set
        End Property

        Public Property ThreadMenuGroupNewThreadVisibility As Visibility
            Get
                Return _threadMenuGroupNewThreadVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_threadMenuGroupNewThreadVisibility, value)
            End Set
        End Property

        Public Property ThreadMenuGroupTopSeparatorVisibility As Visibility
            Get
                Return _threadMenuGroupTopSeparatorVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_threadMenuGroupTopSeparatorVisibility, value)
            End Set
        End Property

        Public Property ThreadMenuGroupOpenVsCodeVisibility As Visibility
            Get
                Return _threadMenuGroupOpenVsCodeVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_threadMenuGroupOpenVsCodeVisibility, value)
            End Set
        End Property

        Public Property ThreadMenuGroupOpenTerminalVisibility As Visibility
            Get
                Return _threadMenuGroupOpenTerminalVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_threadMenuGroupOpenTerminalVisibility, value)
            End Set
        End Property

        Public Property ThreadMenuSelectVisibility As Visibility
            Get
                Return _threadMenuSelectVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_threadMenuSelectVisibility, value)
            End Set
        End Property

        Public Property ThreadMenuSelectRefreshSeparatorVisibility As Visibility
            Get
                Return _threadMenuSelectRefreshSeparatorVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_threadMenuSelectRefreshSeparatorVisibility, value)
            End Set
        End Property

        Public Property ThreadMenuRefreshVisibility As Visibility
            Get
                Return _threadMenuRefreshVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_threadMenuRefreshVisibility, value)
            End Set
        End Property

        Public Property ThreadMenuRefreshActionsSeparatorVisibility As Visibility
            Get
                Return _threadMenuRefreshActionsSeparatorVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_threadMenuRefreshActionsSeparatorVisibility, value)
            End Set
        End Property

        Public Property ThreadMenuForkVisibility As Visibility
            Get
                Return _threadMenuForkVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_threadMenuForkVisibility, value)
            End Set
        End Property

        Public Property ThreadMenuArchiveVisibility As Visibility
            Get
                Return _threadMenuArchiveVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_threadMenuArchiveVisibility, value)
            End Set
        End Property

        Public Property ThreadMenuUnarchiveVisibility As Visibility
            Get
                Return _threadMenuUnarchiveVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_threadMenuUnarchiveVisibility, value)
            End Set
        End Property

        Public Property ThreadMenuGroupNewThreadEnabled As Boolean
            Get
                Return _threadMenuGroupNewThreadEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_threadMenuGroupNewThreadEnabled, value)
            End Set
        End Property

        Public Property ThreadMenuSelectEnabled As Boolean
            Get
                Return _threadMenuSelectEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_threadMenuSelectEnabled, value)
            End Set
        End Property

        Public Property ThreadMenuGroupOpenVsCodeEnabled As Boolean
            Get
                Return _threadMenuGroupOpenVsCodeEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_threadMenuGroupOpenVsCodeEnabled, value)
            End Set
        End Property

        Public Property ThreadMenuGroupOpenTerminalEnabled As Boolean
            Get
                Return _threadMenuGroupOpenTerminalEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_threadMenuGroupOpenTerminalEnabled, value)
            End Set
        End Property

        Public Property ThreadMenuRefreshEnabled As Boolean
            Get
                Return _threadMenuRefreshEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_threadMenuRefreshEnabled, value)
            End Set
        End Property

        Public Property ThreadMenuForkEnabled As Boolean
            Get
                Return _threadMenuForkEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_threadMenuForkEnabled, value)
            End Set
        End Property

        Public Property ThreadMenuArchiveEnabled As Boolean
            Get
                Return _threadMenuArchiveEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_threadMenuArchiveEnabled, value)
            End Set
        End Property

        Public Property ThreadMenuUnarchiveEnabled As Boolean
            Get
                Return _threadMenuUnarchiveEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_threadMenuUnarchiveEnabled, value)
            End Set
        End Property

        Public Property OpenSortMenuCommand As ICommand
            Get
                Return _openSortMenuCommand
            End Get
            Set(value As ICommand)
                SetProperty(_openSortMenuCommand, value)
            End Set
        End Property

        Public Property OpenFilterMenuCommand As ICommand
            Get
                Return _openFilterMenuCommand
            End Get
            Set(value As ICommand)
                SetProperty(_openFilterMenuCommand, value)
            End Set
        End Property

        Public Property SetSortFromMenuCommand As ICommand
            Get
                Return _setSortFromMenuCommand
            End Get
            Set(value As ICommand)
                SetProperty(_setSortFromMenuCommand, value)
            End Set
        End Property

        Public Property ApplyFilterMenuToggleCommand As ICommand
            Get
                Return _applyFilterMenuToggleCommand
            End Get
            Set(value As ICommand)
                SetProperty(_applyFilterMenuToggleCommand, value)
            End Set
        End Property

        Public Property SelectThreadContextCommand As ICommand
            Get
                Return _selectThreadContextCommand
            End Get
            Set(value As ICommand)
                SetProperty(_selectThreadContextCommand, value)
            End Set
        End Property

        Public Property RefreshThreadContextCommand As ICommand
            Get
                Return _refreshThreadContextCommand
            End Get
            Set(value As ICommand)
                SetProperty(_refreshThreadContextCommand, value)
            End Set
        End Property

        Public Property ForkThreadContextCommand As ICommand
            Get
                Return _forkThreadContextCommand
            End Get
            Set(value As ICommand)
                SetProperty(_forkThreadContextCommand, value)
            End Set
        End Property

        Public Property ArchiveThreadContextCommand As ICommand
            Get
                Return _archiveThreadContextCommand
            End Get
            Set(value As ICommand)
                SetProperty(_archiveThreadContextCommand, value)
            End Set
        End Property

        Public Property UnarchiveThreadContextCommand As ICommand
            Get
                Return _unarchiveThreadContextCommand
            End Get
            Set(value As ICommand)
                SetProperty(_unarchiveThreadContextCommand, value)
            End Set
        End Property

        Public Property NewThreadFromGroupContextCommand As ICommand
            Get
                Return _newThreadFromGroupContextCommand
            End Get
            Set(value As ICommand)
                SetProperty(_newThreadFromGroupContextCommand, value)
            End Set
        End Property

        Public Property OpenVsCodeFromGroupContextCommand As ICommand
            Get
                Return _openVsCodeFromGroupContextCommand
            End Get
            Set(value As ICommand)
                SetProperty(_openVsCodeFromGroupContextCommand, value)
            End Set
        End Property

        Public Property OpenTerminalFromGroupContextCommand As ICommand
            Get
                Return _openTerminalFromGroupContextCommand
            End Get
            Set(value As ICommand)
                SetProperty(_openTerminalFromGroupContextCommand, value)
            End Set
        End Property

        Public Sub UpdateInteractionState(session As SessionStateViewModel,
                                          isThreadsLoading As Boolean,
                                          isThreadContentLoading As Boolean)
            Dim isConnected = session IsNot Nothing AndAlso session.IsConnected
            Dim isAuthenticated = session IsNot Nothing AndAlso session.IsAuthenticated
            Dim isConnectedAndAuthenticated = session IsNot Nothing AndAlso session.IsConnectedAndAuthenticated

            IsLoading = isThreadsLoading
            IsThreadContentLoading = isThreadContentLoading
            IsSearchEnabled = isAuthenticated
            IsSortMenuEnabled = isAuthenticated
            IsFilterMenuEnabled = isAuthenticated
            IsThreadListEnabled = isAuthenticated AndAlso Not isThreadsLoading
            CanRunFullRefresh = isConnectedAndAuthenticated AndAlso Not isThreadsLoading AndAlso Not isThreadContentLoading
            CanAutoLoadSelection = isConnectedAndAuthenticated AndAlso Not isThreadsLoading
            CanRunThreadContextActions = isAuthenticated AndAlso Not isThreadsLoading AndAlso Not isThreadContentLoading
        End Sub

        Public Sub ConfigureThreadContextMenuForThread(canRunActions As Boolean,
                                                       isArchived As Boolean)
            ThreadMenuGroupNewThreadVisibility = Visibility.Collapsed
            ThreadMenuGroupOpenVsCodeVisibility = Visibility.Collapsed
            ThreadMenuGroupOpenTerminalVisibility = Visibility.Collapsed
            ThreadMenuGroupTopSeparatorVisibility = Visibility.Collapsed
            ThreadMenuSelectVisibility = Visibility.Visible
            ThreadMenuSelectRefreshSeparatorVisibility = Visibility.Visible
            ThreadMenuRefreshVisibility = Visibility.Visible
            ThreadMenuRefreshActionsSeparatorVisibility = Visibility.Visible
            ThreadMenuForkVisibility = Visibility.Visible
            ThreadMenuArchiveVisibility = If(isArchived, Visibility.Collapsed, Visibility.Visible)
            ThreadMenuUnarchiveVisibility = If(isArchived, Visibility.Visible, Visibility.Collapsed)

            ThreadMenuGroupNewThreadEnabled = False
            ThreadMenuGroupOpenVsCodeEnabled = False
            ThreadMenuGroupOpenTerminalEnabled = False
            ThreadMenuSelectEnabled = True
            ThreadMenuRefreshEnabled = canRunActions
            ThreadMenuForkEnabled = canRunActions
            ThreadMenuArchiveEnabled = canRunActions AndAlso Not isArchived
            ThreadMenuUnarchiveEnabled = canRunActions AndAlso isArchived
        End Sub

        Public Sub ConfigureThreadContextMenuForGroup(canStartHere As Boolean)
            ThreadMenuGroupNewThreadVisibility = Visibility.Visible
            ThreadMenuGroupOpenVsCodeVisibility = Visibility.Visible
            ThreadMenuGroupOpenTerminalVisibility = Visibility.Visible
            ThreadMenuGroupTopSeparatorVisibility = Visibility.Collapsed
            ThreadMenuSelectVisibility = Visibility.Collapsed
            ThreadMenuSelectRefreshSeparatorVisibility = Visibility.Collapsed
            ThreadMenuRefreshVisibility = Visibility.Collapsed
            ThreadMenuRefreshActionsSeparatorVisibility = Visibility.Collapsed
            ThreadMenuForkVisibility = Visibility.Collapsed
            ThreadMenuArchiveVisibility = Visibility.Collapsed
            ThreadMenuUnarchiveVisibility = Visibility.Collapsed

            ThreadMenuGroupNewThreadEnabled = canStartHere
            ThreadMenuGroupOpenVsCodeEnabled = canStartHere
            ThreadMenuGroupOpenTerminalEnabled = canStartHere
            ThreadMenuSelectEnabled = False
            ThreadMenuRefreshEnabled = False
            ThreadMenuForkEnabled = False
            ThreadMenuArchiveEnabled = False
            ThreadMenuUnarchiveEnabled = False
        End Sub

        Public Sub BeginThreadsRefreshState()
            RefreshErrorText = String.Empty
            LastRefreshStartedUtc = DateTimeOffset.UtcNow
        End Sub

        Public Sub CompleteThreadsRefreshState(threadCount As Integer)
            RefreshErrorText = String.Empty
            LastRefreshThreadCount = threadCount
            LastRefreshCompletedUtc = DateTimeOffset.UtcNow
        End Sub

        Public Sub FailThreadsRefreshState(errorMessage As String)
            RefreshErrorText = If(errorMessage, String.Empty)
            LastRefreshCompletedUtc = DateTimeOffset.UtcNow
        End Sub

        Public Sub UpdateThreadListStateText(connected As Boolean,
                                             authenticated As Boolean,
                                             isThreadsLoading As Boolean,
                                             refreshErrorText As String,
                                             totalThreadCount As Integer,
                                             displayCount As Integer,
                                             hasProjectHeaders As Boolean)
            If Not connected Then
                StateText = "Connect to Codex App Server to load threads."
                Return
            End If

            If Not authenticated Then
                StateText = "Authentication required. Sign in to start or load threads."
                Return
            End If

            If isThreadsLoading Then
                StateText = "Loading threads..."
                Return
            End If

            If Not String.IsNullOrWhiteSpace(refreshErrorText) Then
                StateText = $"Error loading threads: {refreshErrorText}"
                Return
            End If

            If totalThreadCount <= 0 Then
                StateText = "No threads found. Start a new thread to begin."
                Return
            End If

            If displayCount <= 0 Then
                If hasProjectHeaders Then
                    StateText = "All project folders are collapsed. Expand a folder to view threads."
                Else
                    StateText = "No threads match the current search/filter."
                End If
                Return
            End If

            StateText = $"Showing {displayCount} of {totalThreadCount} thread(s)."
        End Sub

        Public Function BeginThreadSelectionLoad(threadId As String) As Integer
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If _selectionLoadVersion = Integer.MaxValue Then
                _selectionLoadVersion = 1
            Else
                _selectionLoadVersion += 1
            End If

            RaisePropertyChanged(NameOf(SelectionLoadVersion))
            SetSelectionLoadThreadId(normalizedThreadId)
            IsThreadContentLoading = True
            LastSelectionLoadStartedUtc = DateTimeOffset.UtcNow
            LastSelectionLoadErrorText = String.Empty
            Return _selectionLoadVersion
        End Function

        Public Function IsCurrentThreadSelectionLoad(loadVersion As Integer,
                                                     threadId As String) As Boolean
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            Return loadVersion = _selectionLoadVersion AndAlso
                   StringComparer.Ordinal.Equals(_selectionLoadThreadId, normalizedThreadId)
        End Function

        Public Function TryCompleteThreadSelectionLoad(loadVersion As Integer) As Boolean
            If loadVersion <> _selectionLoadVersion Then
                Return False
            End If

            LastSelectionLoadCompletedUtc = DateTimeOffset.UtcNow
            ClearThreadSelectionLoadState()
            Return True
        End Function

        Public Sub RecordThreadSelectionLoadError(loadVersion As Integer,
                                                  threadId As String,
                                                  errorMessage As String)
            If Not IsCurrentThreadSelectionLoad(loadVersion, threadId) Then
                Return
            End If

            LastSelectionLoadErrorText = If(errorMessage, String.Empty)
        End Sub

        Public Sub CancelThreadSelectionLoadState()
            ClearThreadSelectionLoadState()
        End Sub

        Private Sub ClearThreadSelectionLoadState()
            SetSelectionLoadThreadId(String.Empty)
            IsThreadContentLoading = False
        End Sub

        Private Sub SetSelectionLoadThreadId(value As String)
            Dim normalizedValue = If(value, String.Empty)
            If SetProperty(_selectionLoadThreadId, normalizedValue, NameOf(SelectionLoadThreadId)) Then
                RaisePropertyChanged(NameOf(HasActiveSelectionLoad))
            End If
        End Sub
    End Class
End Namespace
