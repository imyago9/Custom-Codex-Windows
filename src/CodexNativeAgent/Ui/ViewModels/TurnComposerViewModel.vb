Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Globalization
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Media
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
        Private _usageIndicatorsVisibility As Visibility = Visibility.Visible
        Private _rateLimitBarsVisibility As Visibility = Visibility.Collapsed
        Private _rateLimitIndicatorsVisibility As Visibility = Visibility.Collapsed
        Private _regularRateLimitBarsVisibility As Visibility = Visibility.Collapsed
        Private _sparkRateLimitBarsVisibility As Visibility = Visibility.Collapsed
        Private _regularRateLimitIndicatorsVisibility As Visibility = Visibility.Collapsed
        Private _sparkRateLimitIndicatorsVisibility As Visibility = Visibility.Collapsed
        Private _contextUsageVisibility As Visibility = Visibility.Visible
        Private _arePickersExpanded As Boolean = True
        Private _contextUsagePercent As Double
        Private _contextUsagePercentText As String = "--"
        Private _contextUsageTooltipText As String = "Thread context used: --" & Environment.NewLine & "Usage: unavailable"
        Private _contextUsageStrokeDashArray As DoubleCollection = BuildContextUsageStrokeDashArray(0.0R)

        Private _startTurnCommand As AsyncRelayCommand
        Private _steerTurnCommand As AsyncRelayCommand
        Private _interruptTurnCommand As AsyncRelayCommand
        Private ReadOnly _rateLimitBars As New ObservableCollection(Of TurnComposerRateLimitBarViewModel)()
        Private ReadOnly _regularRateLimitBars As New ObservableCollection(Of TurnComposerRateLimitBarViewModel)()
        Private ReadOnly _sparkRateLimitBars As New ObservableCollection(Of TurnComposerRateLimitBarViewModel)()

        Private Const ContextUsageRingRadius As Double = 11.0R
        Private Const ContextUsageRingStrokeThickness As Double = 2.4R

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

        Public Property UsageIndicatorsVisibility As Visibility
            Get
                Return _usageIndicatorsVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_usageIndicatorsVisibility, value)
            End Set
        End Property

        Public ReadOnly Property RateLimitBars As ObservableCollection(Of TurnComposerRateLimitBarViewModel)
            Get
                Return _rateLimitBars
            End Get
        End Property

        Public ReadOnly Property RegularRateLimitBars As ObservableCollection(Of TurnComposerRateLimitBarViewModel)
            Get
                Return _regularRateLimitBars
            End Get
        End Property

        Public ReadOnly Property SparkRateLimitBars As ObservableCollection(Of TurnComposerRateLimitBarViewModel)
            Get
                Return _sparkRateLimitBars
            End Get
        End Property

        Public Property RateLimitBarsVisibility As Visibility
            Get
                Return _rateLimitBarsVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_rateLimitBarsVisibility, value)
            End Set
        End Property

        Public Property RateLimitIndicatorsVisibility As Visibility
            Get
                Return _rateLimitIndicatorsVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_rateLimitIndicatorsVisibility, value)
            End Set
        End Property

        Public Property RegularRateLimitBarsVisibility As Visibility
            Get
                Return _regularRateLimitBarsVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_regularRateLimitBarsVisibility, value)
            End Set
        End Property

        Public Property SparkRateLimitBarsVisibility As Visibility
            Get
                Return _sparkRateLimitBarsVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_sparkRateLimitBarsVisibility, value)
            End Set
        End Property

        Public Property RegularRateLimitIndicatorsVisibility As Visibility
            Get
                Return _regularRateLimitIndicatorsVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_regularRateLimitIndicatorsVisibility, value)
            End Set
        End Property

        Public Property SparkRateLimitIndicatorsVisibility As Visibility
            Get
                Return _sparkRateLimitIndicatorsVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_sparkRateLimitIndicatorsVisibility, value)
            End Set
        End Property

        Public Property ContextUsageVisibility As Visibility
            Get
                Return _contextUsageVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_contextUsageVisibility, value)
            End Set
        End Property

        Public Property ArePickersExpanded As Boolean
            Get
                Return _arePickersExpanded
            End Get
            Set(value As Boolean)
                If SetProperty(_arePickersExpanded, value) Then
                    SyncUsageIndicatorsVisibility()
                End If
            End Set
        End Property

        Public Property ContextUsagePercent As Double
            Get
                Return _contextUsagePercent
            End Get
            Set(value As Double)
                SetProperty(_contextUsagePercent, value)
            End Set
        End Property

        Public Property ContextUsagePercentText As String
            Get
                Return _contextUsagePercentText
            End Get
            Set(value As String)
                SetProperty(_contextUsagePercentText, If(value, String.Empty))
            End Set
        End Property

        Public Property ContextUsageTooltipText As String
            Get
                Return _contextUsageTooltipText
            End Get
            Set(value As String)
                SetProperty(_contextUsageTooltipText, If(value, String.Empty))
            End Set
        End Property

        Public Property ContextUsageStrokeDashArray As DoubleCollection
            Get
                Return _contextUsageStrokeDashArray
            End Get
            Set(value As DoubleCollection)
                SetProperty(_contextUsageStrokeDashArray, If(value, BuildContextUsageStrokeDashArray(0.0R)))
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

        Public Sub SetRateLimitBars(items As IEnumerable(Of TurnComposerRateLimitBarViewModel))
            _rateLimitBars.Clear()
            _regularRateLimitBars.Clear()
            _sparkRateLimitBars.Clear()
            If items IsNot Nothing Then
                For Each item In items
                    If item Is Nothing Then
                        Continue For
                    End If

                    _rateLimitBars.Add(item)
                    Dim groupKey = NormalizeRateLimitGroupKey(item.GroupKey)
                    If StringComparer.Ordinal.Equals(groupKey, "spark") Then
                        _sparkRateLimitBars.Add(item)
                    Else
                        _regularRateLimitBars.Add(item)
                    End If
                Next
            End If

            RateLimitBarsVisibility = If(_rateLimitBars.Count > 0, Visibility.Visible, Visibility.Collapsed)
            RegularRateLimitBarsVisibility = If(_regularRateLimitBars.Count > 0, Visibility.Visible, Visibility.Collapsed)
            SparkRateLimitBarsVisibility = If(_sparkRateLimitBars.Count > 0, Visibility.Visible, Visibility.Collapsed)
            SyncUsageIndicatorsVisibility()
        End Sub

        Public Sub SetContextUsageIndicator(contextPercent As Double?, tooltipText As String)
            If contextPercent.HasValue Then
                Dim normalizedPercent = ClampPercent(contextPercent.Value)
                ContextUsagePercent = normalizedPercent
                ContextUsagePercentText = Math.Round(normalizedPercent,
                                                     MidpointRounding.AwayFromZero).ToString("0",
                                                                                              CultureInfo.InvariantCulture) & "%"
                Dim normalizedTooltip = If(tooltipText, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalizedTooltip) Then
                    normalizedTooltip = $"Thread context used: {normalizedPercent.ToString("0.#", CultureInfo.InvariantCulture)}%{Environment.NewLine}Usage: unavailable"
                End If
                ContextUsageTooltipText = normalizedTooltip
                ContextUsageStrokeDashArray = BuildContextUsageStrokeDashArray(normalizedPercent)
                ContextUsageVisibility = Visibility.Visible
            Else
                ContextUsagePercent = 0.0R
                ContextUsagePercentText = "--"
                ContextUsageTooltipText = "Thread context used: --" & Environment.NewLine & "Usage: unavailable"
                ContextUsageStrokeDashArray = BuildContextUsageStrokeDashArray(0.0R)
                ContextUsageVisibility = Visibility.Visible
            End If

            SyncUsageIndicatorsVisibility()
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

        Private Sub SyncUsageIndicatorsVisibility()
            Dim hasRegularRateLimitBars = RegularRateLimitBarsVisibility = Visibility.Visible
            Dim hasSparkRateLimitBars = SparkRateLimitBarsVisibility = Visibility.Visible
            Dim hasContextUsage = ContextUsageVisibility = Visibility.Visible
            RegularRateLimitIndicatorsVisibility = If(hasRegularRateLimitBars,
                                                      Visibility.Visible,
                                                      Visibility.Collapsed)
            SparkRateLimitIndicatorsVisibility = If(hasSparkRateLimitBars,
                                                    Visibility.Visible,
                                                    Visibility.Collapsed)
            RateLimitIndicatorsVisibility = If(RegularRateLimitIndicatorsVisibility = Visibility.Visible OrElse
                                               SparkRateLimitIndicatorsVisibility = Visibility.Visible,
                                               Visibility.Visible,
                                               Visibility.Collapsed)
            UsageIndicatorsVisibility = If(RateLimitIndicatorsVisibility = Visibility.Visible OrElse hasContextUsage,
                                           Visibility.Visible,
                                           Visibility.Collapsed)
        End Sub

        Private Shared Function NormalizeRateLimitGroupKey(groupKey As String) As String
            Dim normalized = If(groupKey, String.Empty).Trim()
            If normalized.Length = 0 Then
                Return "regular"
            End If

            If StringComparer.OrdinalIgnoreCase.Equals(normalized, "spark") Then
                Return "spark"
            End If

            Return "regular"
        End Function

        Private Shared Function BuildContextUsageStrokeDashArray(percent As Double) As DoubleCollection
            Dim normalizedPercent = ClampPercent(percent)
            Dim normalizedRatio = normalizedPercent / 100.0R
            Dim circumference = 2.0R * Math.PI * ContextUsageRingRadius
            Dim totalDashUnits = circumference / ContextUsageRingStrokeThickness
            Dim dashUnits = totalDashUnits * normalizedRatio
            Dim gapUnits = Math.Max(0.001R, totalDashUnits - dashUnits)

            Return New DoubleCollection() From {
                Math.Max(0.0R, dashUnits),
                gapUnits
            }
        End Function

        Private Shared Function ClampPercent(value As Double) As Double
            If Double.IsNaN(value) OrElse Double.IsInfinity(value) Then
                Return 0.0R
            End If

            If value < 0.0R Then
                Return 0.0R
            End If

            If value > 100.0R Then
                Return 100.0R
            End If

            Return value
        End Function
    End Class

    Public NotInheritable Class TurnComposerRateLimitBarViewModel
        Private Const RingRadius As Double = 9.0R
        Private Const RingStrokeThickness As Double = 2.0R
        Private _remainingPercent As Double
        Private _remainingStrokeDashArray As DoubleCollection = BuildRemainingStrokeDashArray(0.0R)

        Public Property BarId As String = String.Empty
        Public Property GroupKey As String = "regular"
        Public Property ShortLabel As String = "--"

        Public Property RemainingPercent As Double
            Get
                Return _remainingPercent
            End Get
            Set(value As Double)
                _remainingPercent = ClampPercent(value)
                _remainingStrokeDashArray = BuildRemainingStrokeDashArray(_remainingPercent)
            End Set
        End Property

        Public Property RemainingStrokeDashArray As DoubleCollection
            Get
                Return _remainingStrokeDashArray
            End Get
            Set(value As DoubleCollection)
                _remainingStrokeDashArray = If(value, BuildRemainingStrokeDashArray(_remainingPercent))
            End Set
        End Property

        Public Property UsedPercent As Double
        Public Property TooltipText As String = String.Empty
        Public Property BarBrush As Brush = Brushes.Transparent

        Private Shared Function BuildRemainingStrokeDashArray(percent As Double) As DoubleCollection
            Dim normalizedPercent = ClampPercent(percent)
            Dim normalizedRatio = normalizedPercent / 100.0R
            Dim circumference = 2.0R * Math.PI * RingRadius
            Dim totalDashUnits = circumference / RingStrokeThickness
            Dim dashUnits = totalDashUnits * normalizedRatio
            Dim gapUnits = Math.Max(0.001R, totalDashUnits - dashUnits)

            Return New DoubleCollection() From {
                Math.Max(0.0R, dashUnits),
                gapUnits
            }
        End Function

        Private Shared Function ClampPercent(value As Double) As Double
            If Double.IsNaN(value) OrElse Double.IsInfinity(value) Then
                Return 0.0R
            End If

            If value < 0.0R Then
                Return 0.0R
            End If

            If value > 100.0R Then
                Return 100.0R
            End If

            Return value
        End Function
    End Class
End Namespace
