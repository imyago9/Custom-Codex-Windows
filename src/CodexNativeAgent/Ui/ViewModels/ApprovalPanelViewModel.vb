Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Input
Imports CodexNativeAgent.Ui.Mvvm

Namespace CodexNativeAgent.Ui.ViewModels
    Public NotInheritable Class ApprovalPanelViewModel
        Inherits ViewModelBase

        Private _summaryText As String = String.Empty
        Private _cardVisibility As Visibility = Visibility.Collapsed
        Private _canAccept As Boolean
        Private _canAcceptSession As Boolean
        Private _canAcceptAmended As Boolean
        Private _acceptAmendedVisibility As Visibility = Visibility.Collapsed
        Private _canDecline As Boolean
        Private _canCancel As Boolean
        Private _pendingQueueCount As Integer
        Private _activeMethodName As String = String.Empty
        Private _supportsExecpolicyAmendment As Boolean
        Private _selectedOptionNumber As Integer = 1
        Private _lastQueuedMethodName As String = String.Empty
        Private _lastResolvedAction As String = String.Empty
        Private _lastResolvedDecision As String = String.Empty
        Private _lastErrorText As String = String.Empty
        Private _lastQueueUpdatedUtc As DateTimeOffset?
        Private _lastResolvedUtc As DateTimeOffset?

        Private _acceptCommand As ICommand
        Private _acceptSessionCommand As ICommand
        Private _acceptAmendedCommand As ICommand
        Private _declineCommand As ICommand
        Private _cancelCommand As ICommand

        Public Property SummaryText As String
            Get
                Return _summaryText
            End Get
            Set(value As String)
                SetProperty(_summaryText, If(value, String.Empty))
            End Set
        End Property

        Public Property CardVisibility As Visibility
            Get
                Return _cardVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_cardVisibility, value)
            End Set
        End Property

        Public Property CanAccept As Boolean
            Get
                Return _canAccept
            End Get
            Set(value As Boolean)
                SetProperty(_canAccept, value)
            End Set
        End Property

        Public Property CanAcceptSession As Boolean
            Get
                Return _canAcceptSession
            End Get
            Set(value As Boolean)
                SetProperty(_canAcceptSession, value)
            End Set
        End Property

        Public Property CanAcceptAmended As Boolean
            Get
                Return _canAcceptAmended
            End Get
            Set(value As Boolean)
                SetProperty(_canAcceptAmended, value)
            End Set
        End Property

        Public Property AcceptAmendedVisibility As Visibility
            Get
                Return _acceptAmendedVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_acceptAmendedVisibility, value)
            End Set
        End Property

        Public Property CanDecline As Boolean
            Get
                Return _canDecline
            End Get
            Set(value As Boolean)
                SetProperty(_canDecline, value)
            End Set
        End Property

        Public Property CanCancel As Boolean
            Get
                Return _canCancel
            End Get
            Set(value As Boolean)
                SetProperty(_canCancel, value)
            End Set
        End Property

        Public Property PendingQueueCount As Integer
            Get
                Return _pendingQueueCount
            End Get
            Set(value As Integer)
                SetProperty(_pendingQueueCount, Math.Max(0, value))
            End Set
        End Property

        Public Property SelectedOptionNumber As Integer
            Get
                Return _selectedOptionNumber
            End Get
            Set(value As Integer)
                Dim normalized = Math.Max(1, Math.Min(5, value))
                SetProperty(_selectedOptionNumber, normalized)
            End Set
        End Property

        Public Property ActiveMethodName As String
            Get
                Return _activeMethodName
            End Get
            Set(value As String)
                SetProperty(_activeMethodName, If(value, String.Empty))
            End Set
        End Property

        Public Property LastQueuedMethodName As String
            Get
                Return _lastQueuedMethodName
            End Get
            Set(value As String)
                SetProperty(_lastQueuedMethodName, If(value, String.Empty))
            End Set
        End Property

        Public Property LastResolvedAction As String
            Get
                Return _lastResolvedAction
            End Get
            Set(value As String)
                SetProperty(_lastResolvedAction, If(value, String.Empty))
            End Set
        End Property

        Public Property LastResolvedDecision As String
            Get
                Return _lastResolvedDecision
            End Get
            Set(value As String)
                SetProperty(_lastResolvedDecision, If(value, String.Empty))
            End Set
        End Property

        Public Property LastErrorText As String
            Get
                Return _lastErrorText
            End Get
            Set(value As String)
                SetProperty(_lastErrorText, If(value, String.Empty))
            End Set
        End Property

        Public Property LastQueueUpdatedUtc As DateTimeOffset?
            Get
                Return _lastQueueUpdatedUtc
            End Get
            Set(value As DateTimeOffset?)
                SetProperty(_lastQueueUpdatedUtc, value)
            End Set
        End Property

        Public Property LastResolvedUtc As DateTimeOffset?
            Get
                Return _lastResolvedUtc
            End Get
            Set(value As DateTimeOffset?)
                SetProperty(_lastResolvedUtc, value)
            End Set
        End Property

        Public Property AcceptCommand As ICommand
            Get
                Return _acceptCommand
            End Get
            Set(value As ICommand)
                SetProperty(_acceptCommand, value)
            End Set
        End Property

        Public Property AcceptSessionCommand As ICommand
            Get
                Return _acceptSessionCommand
            End Get
            Set(value As ICommand)
                SetProperty(_acceptSessionCommand, value)
            End Set
        End Property

        Public Property AcceptAmendedCommand As ICommand
            Get
                Return _acceptAmendedCommand
            End Get
            Set(value As ICommand)
                SetProperty(_acceptAmendedCommand, value)
            End Set
        End Property

        Public Property DeclineCommand As ICommand
            Get
                Return _declineCommand
            End Get
            Set(value As ICommand)
                SetProperty(_declineCommand, value)
            End Set
        End Property

        Public Property CancelCommand As ICommand
            Get
                Return _cancelCommand
            End Get
            Set(value As ICommand)
                SetProperty(_cancelCommand, value)
            End Set
        End Property

        Public Sub ConfigureCommands(acceptAsync As Func(Of Task),
                                     acceptSessionAsync As Func(Of Task),
                                     acceptAmendedAsync As Func(Of Task),
                                     declineAsync As Func(Of Task),
                                     cancelAsync As Func(Of Task))
            AcceptCommand = New AsyncRelayCommand(acceptAsync)
            AcceptSessionCommand = New AsyncRelayCommand(acceptSessionAsync)
            AcceptAmendedCommand = New AsyncRelayCommand(acceptAmendedAsync)
            DeclineCommand = New AsyncRelayCommand(declineAsync)
            CancelCommand = New AsyncRelayCommand(cancelAsync)
        End Sub

        Public Sub UpdateAvailability(isAuthenticated As Boolean,
                                      hasActiveApproval As Boolean)
            CanAccept = isAuthenticated AndAlso hasActiveApproval
            CanAcceptSession = isAuthenticated AndAlso hasActiveApproval
            CanAcceptAmended = isAuthenticated AndAlso hasActiveApproval AndAlso _supportsExecpolicyAmendment
            AcceptAmendedVisibility = If(hasActiveApproval AndAlso _supportsExecpolicyAmendment, Visibility.Visible, Visibility.Collapsed)
            CanDecline = isAuthenticated AndAlso hasActiveApproval
            CanCancel = isAuthenticated AndAlso hasActiveApproval
            CardVisibility = If(hasActiveApproval, Visibility.Visible, Visibility.Collapsed)
            EnsureSelectedOptionIsValid()
        End Sub

        Public Sub SetThreadScopedState(summaryText As String,
                                        activeMethodName As String,
                                        supportsExecpolicyAmendment As Boolean,
                                        pendingQueueCount As Integer)
            SummaryText = If(summaryText, String.Empty)
            ActiveMethodName = If(activeMethodName, String.Empty)
            _supportsExecpolicyAmendment = supportsExecpolicyAmendment
            PendingQueueCount = Math.Max(0, pendingQueueCount)
        End Sub

        Public Sub ResetLifecycleState()
            SummaryText = String.Empty
            PendingQueueCount = 0
            ActiveMethodName = String.Empty
            _supportsExecpolicyAmendment = False
            LastQueuedMethodName = String.Empty
            LastResolvedAction = String.Empty
            LastResolvedDecision = String.Empty
            LastErrorText = String.Empty
            LastQueueUpdatedUtc = Nothing
            LastResolvedUtc = Nothing
            SelectedOptionNumber = 1
            UpdateAvailability(False, False)
        End Sub

        Public Sub OnApprovalQueued(methodName As String,
                                    pendingQueueCount As Integer)
            LastQueuedMethodName = If(methodName, String.Empty)
            PendingQueueCount = pendingQueueCount
            LastQueueUpdatedUtc = DateTimeOffset.UtcNow
            LastErrorText = String.Empty
        End Sub

        Public Sub OnApprovalActivated(methodName As String,
                                       summary As String,
                                       supportsExecpolicyAmendment As Boolean,
                                       pendingQueueCount As Integer)
            ActiveMethodName = If(methodName, String.Empty)
            SummaryText = If(summary, String.Empty)
            _supportsExecpolicyAmendment = supportsExecpolicyAmendment
            PendingQueueCount = pendingQueueCount
            LastQueueUpdatedUtc = DateTimeOffset.UtcNow
            LastErrorText = String.Empty
        End Sub

        Public Sub OnApprovalQueueEmpty()
            ActiveMethodName = String.Empty
            SummaryText = String.Empty
            _supportsExecpolicyAmendment = False
            PendingQueueCount = 0
            LastQueueUpdatedUtc = DateTimeOffset.UtcNow
        End Sub

        Public Sub OnApprovalResolved(action As String,
                                      decision As String,
                                      pendingQueueCount As Integer)
            LastResolvedAction = If(action, String.Empty)
            LastResolvedDecision = If(decision, String.Empty)
            LastResolvedUtc = DateTimeOffset.UtcNow
            PendingQueueCount = Math.Max(0, pendingQueueCount)
            LastErrorText = String.Empty
            ActiveMethodName = String.Empty
            _supportsExecpolicyAmendment = False
        End Sub

        Public Sub RecordError(errorMessage As String)
            LastErrorText = If(errorMessage, String.Empty)
        End Sub

        Public Function MoveSelection(delta As Integer) As Boolean
            If CardVisibility <> Visibility.Visible Then
                Return False
            End If

            Dim selectableOptions = BuildSelectableOptionList()
            If selectableOptions.Count = 0 Then
                Return False
            End If

            Dim currentIndex = selectableOptions.IndexOf(SelectedOptionNumber)
            If currentIndex < 0 Then
                currentIndex = 0
            End If

            Dim stepDirection = If(delta < 0, -1, 1)
            Dim nextIndex = (currentIndex + stepDirection + selectableOptions.Count) Mod selectableOptions.Count
            SelectedOptionNumber = selectableOptions(nextIndex)
            Return True
        End Function

        Public Function TryExecuteSelectedOption() As Boolean
            Return TryExecuteOption(SelectedOptionNumber)
        End Function

        Public Function TryExecuteOption(optionNumber As Integer) As Boolean
            If CardVisibility <> Visibility.Visible Then
                Return False
            End If

            Dim normalizedOption = Math.Max(1, Math.Min(5, optionNumber))
            SelectedOptionNumber = normalizedOption
            Select Case normalizedOption
                Case 1
                    Return TryExecuteOptionCommand(normalizedOption, AcceptCommand, CanAccept)
                Case 2
                    Return TryExecuteOptionCommand(normalizedOption, AcceptSessionCommand, CanAcceptSession)
                Case 3
                    Return TryExecuteOptionCommand(normalizedOption,
                                                   AcceptAmendedCommand,
                                                   CanAcceptAmended AndAlso AcceptAmendedVisibility = Visibility.Visible)
                Case 4
                    Return TryExecuteOptionCommand(normalizedOption, DeclineCommand, CanDecline)
                Case 5
                    Return TryExecuteOptionCommand(normalizedOption, CancelCommand, CanCancel)
                Case Else
                    Return False
            End Select
        End Function

        Private Function TryExecuteOptionCommand(optionNumber As Integer,
                                                 command As ICommand,
                                                 isAvailable As Boolean) As Boolean
            If Not isAvailable OrElse command Is Nothing OrElse Not command.CanExecute(Nothing) Then
                Return False
            End If

            SelectedOptionNumber = optionNumber
            command.Execute(Nothing)
            Return True
        End Function

        Private Sub EnsureSelectedOptionIsValid()
            Dim selectableOptions = BuildSelectableOptionList()
            If selectableOptions.Count = 0 Then
                SelectedOptionNumber = 1
                Return
            End If

            If selectableOptions.Contains(SelectedOptionNumber) Then
                Return
            End If

            SelectedOptionNumber = selectableOptions(0)
        End Sub

        Private Function BuildSelectableOptionList() As List(Of Integer)
            Dim options As New List(Of Integer)()
            If CanAccept Then
                options.Add(1)
            End If
            If CanAcceptSession Then
                options.Add(2)
            End If
            If CanAcceptAmended AndAlso AcceptAmendedVisibility = Visibility.Visible Then
                options.Add(3)
            End If
            If CanDecline Then
                options.Add(4)
            End If
            If CanCancel Then
                options.Add(5)
            End If

            Return options
        End Function
    End Class
End Namespace
