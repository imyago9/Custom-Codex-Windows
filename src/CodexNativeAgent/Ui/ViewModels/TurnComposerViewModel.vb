Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Input
Imports CodexNativeAgent.Ui.Mvvm

Namespace CodexNativeAgent.Ui.ViewModels
    Public NotInheritable Class TurnComposerViewModel
        Inherits ViewModelBase

        Private _inputText As String = String.Empty
        Private _selectedModelId As String = String.Empty
        Private _selectedReasoningEffort As String = "medium"
        Private _selectedApprovalPolicy As String = "on-request"
        Private _selectedSandbox As String = "workspace-write"

        Private _isInputEnabled As Boolean
        Private _isModelEnabled As Boolean
        Private _isReasoningEnabled As Boolean
        Private _isApprovalPolicyEnabled As Boolean
        Private _isSandboxEnabled As Boolean

        Private _canStartTurn As Boolean
        Private _canSteerTurn As Boolean
        Private _canInterruptTurn As Boolean
        Private _startTurnVisibility As Visibility = Visibility.Visible
        Private _interruptTurnVisibility As Visibility = Visibility.Collapsed

        Private _startTurnCommand As AsyncRelayCommand
        Private _steerTurnCommand As AsyncRelayCommand
        Private _interruptTurnCommand As AsyncRelayCommand

        Public Sub New()
            _startTurnCommand = New AsyncRelayCommand(Function() Task.CompletedTask, Function() False)
            _steerTurnCommand = New AsyncRelayCommand(Function() Task.CompletedTask, Function() False)
            _interruptTurnCommand = New AsyncRelayCommand(Function() Task.CompletedTask, Function() False)
        End Sub

        Public Property InputText As String
            Get
                Return _inputText
            End Get
            Set(value As String)
                SetProperty(_inputText, If(value, String.Empty))
            End Set
        End Property

        Public Property SelectedModelId As String
            Get
                Return _selectedModelId
            End Get
            Set(value As String)
                SetProperty(_selectedModelId, If(value, String.Empty))
            End Set
        End Property

        Public Property SelectedReasoningEffort As String
            Get
                Return _selectedReasoningEffort
            End Get
            Set(value As String)
                SetProperty(_selectedReasoningEffort, If(value, String.Empty))
            End Set
        End Property

        Public Property SelectedApprovalPolicy As String
            Get
                Return _selectedApprovalPolicy
            End Get
            Set(value As String)
                SetProperty(_selectedApprovalPolicy, If(value, String.Empty))
            End Set
        End Property

        Public Property SelectedSandbox As String
            Get
                Return _selectedSandbox
            End Get
            Set(value As String)
                SetProperty(_selectedSandbox, If(value, String.Empty))
            End Set
        End Property

        Public Property IsInputEnabled As Boolean
            Get
                Return _isInputEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isInputEnabled, value)
            End Set
        End Property

        Public Property IsModelEnabled As Boolean
            Get
                Return _isModelEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isModelEnabled, value)
            End Set
        End Property

        Public Property IsReasoningEnabled As Boolean
            Get
                Return _isReasoningEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isReasoningEnabled, value)
            End Set
        End Property

        Public Property IsApprovalPolicyEnabled As Boolean
            Get
                Return _isApprovalPolicyEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isApprovalPolicyEnabled, value)
            End Set
        End Property

        Public Property IsSandboxEnabled As Boolean
            Get
                Return _isSandboxEnabled
            End Get
            Set(value As Boolean)
                SetProperty(_isSandboxEnabled, value)
            End Set
        End Property

        Public Property CanStartTurn As Boolean
            Get
                Return _canStartTurn
            End Get
            Set(value As Boolean)
                If SetProperty(_canStartTurn, value) Then
                    RaiseTurnCommandCanExecuteChanged()
                End If
            End Set
        End Property

        Public Property CanSteerTurn As Boolean
            Get
                Return _canSteerTurn
            End Get
            Set(value As Boolean)
                If SetProperty(_canSteerTurn, value) Then
                    RaiseTurnCommandCanExecuteChanged()
                End If
            End Set
        End Property

        Public Property CanInterruptTurn As Boolean
            Get
                Return _canInterruptTurn
            End Get
            Set(value As Boolean)
                If SetProperty(_canInterruptTurn, value) Then
                    RaiseTurnCommandCanExecuteChanged()
                End If
            End Set
        End Property

        Public Property StartTurnVisibility As Visibility
            Get
                Return _startTurnVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_startTurnVisibility, value)
            End Set
        End Property

        Public Property InterruptTurnVisibility As Visibility
            Get
                Return _interruptTurnVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_interruptTurnVisibility, value)
            End Set
        End Property

        Public ReadOnly Property StartTurnCommand As ICommand
            Get
                Return _startTurnCommand
            End Get
        End Property

        Public ReadOnly Property SteerTurnCommand As ICommand
            Get
                Return _steerTurnCommand
            End Get
        End Property

        Public ReadOnly Property InterruptTurnCommand As ICommand
            Get
                Return _interruptTurnCommand
            End Get
        End Property

        Public Sub ConfigureCommands(startTurnAsync As Func(Of Task),
                                     steerTurnAsync As Func(Of Task),
                                     interruptTurnAsync As Func(Of Task))
            _startTurnCommand = New AsyncRelayCommand(startTurnAsync, Function() CanStartTurn)
            _steerTurnCommand = New AsyncRelayCommand(steerTurnAsync, Function() CanSteerTurn)
            _interruptTurnCommand = New AsyncRelayCommand(interruptTurnAsync, Function() CanInterruptTurn)

            RaisePropertyChanged(NameOf(StartTurnCommand))
            RaisePropertyChanged(NameOf(SteerTurnCommand))
            RaisePropertyChanged(NameOf(InterruptTurnCommand))
            RaiseTurnCommandCanExecuteChanged()
        End Sub

        Private Sub RaiseTurnCommandCanExecuteChanged()
            If _startTurnCommand IsNot Nothing Then
                _startTurnCommand.RaiseCanExecuteChanged()
            End If

            If _steerTurnCommand IsNot Nothing Then
                _steerTurnCommand.RaiseCanExecuteChanged()
            End If

            If _interruptTurnCommand IsNot Nothing Then
                _interruptTurnCommand.RaiseCanExecuteChanged()
            End If
        End Sub
    End Class
End Namespace
