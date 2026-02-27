Imports System.Collections.Generic
Imports System.ComponentModel
Imports CodexNativeAgent.Ui.ViewModels

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private Const GlobalTurnComposerStateKey As String = "__global_turn_composer__"

        Private NotInheritable Class TurnComposerThreadState
            Public Property InputText As String = String.Empty
            Public Property ModelId As String = String.Empty
            Public Property ReasoningEffort As String = "medium"
            Public Property ApprovalPolicy As String = "on-request"
            Public Property Sandbox As String = "workspace-write"
        End Class

        Private ReadOnly _turnComposerStatesByThreadKey As New Dictionary(Of String, TurnComposerThreadState)(StringComparer.Ordinal)
        Private _suppressTurnComposerStateSync As Boolean

        Private Sub OnTurnComposerPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            If _suppressTurnComposerStateSync Then
                Return
            End If

            If e Is Nothing OrElse String.IsNullOrWhiteSpace(e.PropertyName) Then
                SyncTurnComposerStateForCurrentSelection()
                Return
            End If

            Select Case e.PropertyName
                Case NameOf(TurnComposerViewModel.InputText),
                     NameOf(TurnComposerViewModel.SelectedModelId),
                     NameOf(TurnComposerViewModel.SelectedReasoningEffort),
                     NameOf(TurnComposerViewModel.SelectedApprovalPolicy),
                     NameOf(TurnComposerViewModel.SelectedSandbox)
                    SyncTurnComposerStateForCurrentSelection()
            End Select
        End Sub

        Private Sub ClearTurnComposerThreadStates()
            _turnComposerStatesByThreadKey.Clear()
        End Sub

        Private Sub SyncTurnComposerStateForCurrentSelection()
            Dim stateKey = ResolveTurnComposerStateKeyForCurrentSelection()
            If String.IsNullOrWhiteSpace(stateKey) Then
                Return
            End If

            _turnComposerStatesByThreadKey(stateKey) = CreateComposerStateFromCurrentComposer(preserveInputText:=True)
        End Sub

        Private Sub RestoreTurnComposerStateForCurrentSelection(Optional reason As String = Nothing)
            Dim stateKey = ResolveTurnComposerStateKeyForCurrentSelection()
            If String.IsNullOrWhiteSpace(stateKey) Then
                Return
            End If

            Dim state = CloneTurnComposerState(GetOrCreateTurnComposerState(stateKey))
            Dim normalizedState = NormalizeTurnComposerStateForCurrentUi(state)
            _turnComposerStatesByThreadKey(stateKey) = CloneTurnComposerState(normalizedState)
            ApplyTurnComposerStateToComposer(normalizedState)
        End Sub

        Private Function ResolveTurnComposerStateForThread(threadId As String,
                                                           allowDraftWhenNoThread As Boolean) As TurnComposerThreadState
            Dim stateKey = ResolveTurnComposerStateKey(threadId,
                                                       allowDraftWhenNoThread:=allowDraftWhenNoThread,
                                                       allowGlobalFallback:=True)
            If String.IsNullOrWhiteSpace(stateKey) Then
                Return CreateComposerStateFromCurrentComposer(preserveInputText:=True)
            End If

            Dim state = CloneTurnComposerState(GetOrCreateTurnComposerState(stateKey))
            Dim normalizedState = NormalizeTurnComposerStateForCurrentUi(state)
            _turnComposerStatesByThreadKey(stateKey) = CloneTurnComposerState(normalizedState)
            Return normalizedState
        End Function

        Private Sub PromotePendingDraftTurnComposerStateToThread(threadId As String)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            Dim pendingState As TurnComposerThreadState = Nothing
            If Not _turnComposerStatesByThreadKey.TryGetValue(PendingNewThreadTranscriptTabId, pendingState) OrElse pendingState Is Nothing Then
                Return
            End If

            Dim promotedState = NormalizeTurnComposerStateForCurrentUi(CloneTurnComposerState(pendingState))
            _turnComposerStatesByThreadKey(normalizedThreadId) = promotedState

            Dim pendingNextState = CloneTurnComposerState(pendingState)
            pendingNextState.InputText = String.Empty
            _turnComposerStatesByThreadKey(PendingNewThreadTranscriptTabId) = NormalizeTurnComposerStateForCurrentUi(pendingNextState)
        End Sub

        Private Function ResolveTurnComposerStateKeyForCurrentSelection() As String
            Return ResolveTurnComposerStateKey(GetVisibleThreadId(),
                                               allowDraftWhenNoThread:=True,
                                               allowGlobalFallback:=True)
        End Function

        Private Function ResolveTurnComposerStateKey(threadId As String,
                                                     allowDraftWhenNoThread As Boolean,
                                                     allowGlobalFallback As Boolean) As String
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return normalizedThreadId
            End If

            If allowDraftWhenNoThread Then
                Dim pendingDraftSelectionActive = _pendingNewThreadFirstPromptSelection
                If Not pendingDraftSelectionActive Then
                    pendingDraftSelectionActive = IsPendingNewThreadTranscriptTabActive()
                End If

                If pendingDraftSelectionActive Then
                    Return PendingNewThreadTranscriptTabId
                End If
            End If

            If allowGlobalFallback Then
                Return GlobalTurnComposerStateKey
            End If

            Return String.Empty
        End Function

        Private Function GetOrCreateTurnComposerState(stateKey As String) As TurnComposerThreadState
            Dim normalizedStateKey = If(stateKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedStateKey) Then
                Return CreateComposerStateFromCurrentComposer(preserveInputText:=True)
            End If

            Dim existingState As TurnComposerThreadState = Nothing
            If _turnComposerStatesByThreadKey.TryGetValue(normalizedStateKey, existingState) AndAlso existingState IsNot Nothing Then
                Return existingState
            End If

            Dim createdState = CreateComposerStateFromCurrentComposer(preserveInputText:=False)
            _turnComposerStatesByThreadKey(normalizedStateKey) = createdState
            Return createdState
        End Function

        Private Function CreateComposerStateFromCurrentComposer(preserveInputText As Boolean) As TurnComposerThreadState
            Dim state As New TurnComposerThreadState()
            If _viewModel Is Nothing OrElse _viewModel.TurnComposer Is Nothing Then
                Return NormalizeTurnComposerStateForCurrentUi(state)
            End If

            state.InputText = If(preserveInputText,
                                 If(_viewModel.TurnComposer.InputText, String.Empty),
                                 String.Empty)
            state.ModelId = If(_viewModel.TurnComposer.SelectedModelId, String.Empty)
            state.ReasoningEffort = If(_viewModel.TurnComposer.SelectedReasoningEffort, String.Empty)
            state.ApprovalPolicy = If(_viewModel.TurnComposer.SelectedApprovalPolicy, String.Empty)
            state.Sandbox = If(_viewModel.TurnComposer.SelectedSandbox, String.Empty)
            Return NormalizeTurnComposerStateForCurrentUi(state)
        End Function

        Private Sub ApplyTurnComposerStateToComposer(state As TurnComposerThreadState)
            If state Is Nothing OrElse _viewModel Is Nothing OrElse _viewModel.TurnComposer Is Nothing Then
                Return
            End If

            _suppressTurnComposerStateSync = True
            Try
                _viewModel.TurnComposer.InputText = If(state.InputText, String.Empty)
                _viewModel.TurnComposer.SelectedModelId = If(state.ModelId, String.Empty)
                _viewModel.TurnComposer.SelectedReasoningEffort = If(state.ReasoningEffort, String.Empty)
                _viewModel.TurnComposer.SelectedApprovalPolicy = If(state.ApprovalPolicy, String.Empty)
                _viewModel.TurnComposer.SelectedSandbox = If(state.Sandbox, String.Empty)
            Finally
                _suppressTurnComposerStateSync = False
            End Try
        End Sub

        Private Function NormalizeTurnComposerStateForCurrentUi(state As TurnComposerThreadState) As TurnComposerThreadState
            Dim source = If(state, New TurnComposerThreadState())
            Dim normalized As New TurnComposerThreadState() With {
                .InputText = If(source.InputText, String.Empty),
                .ModelId = ResolveAvailableModelId(source.ModelId),
                .ReasoningEffort = NormalizeReasoningEffort(source.ReasoningEffort),
                .ApprovalPolicy = NormalizeApprovalPolicy(source.ApprovalPolicy),
                .Sandbox = NormalizeSandboxMode(source.Sandbox)
            }
            Return normalized
        End Function

        Private Function ResolveAvailableModelId(preferredModelId As String) As String
            Dim normalizedPreferredModelId = If(preferredModelId, String.Empty).Trim()
            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.CmbModel Is Nothing Then
                Return normalizedPreferredModelId
            End If

            Dim firstModelId As String = String.Empty
            Dim defaultModelId As String = String.Empty
            For Each itemObject In WorkspacePaneHost.CmbModel.Items
                Dim modelEntry = TryCast(itemObject, ModelListEntry)
                If modelEntry Is Nothing Then
                    Continue For
                End If

                Dim modelId = If(modelEntry.Id, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(modelId) Then
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(firstModelId) Then
                    firstModelId = modelId
                End If

                If modelEntry.IsDefault AndAlso String.IsNullOrWhiteSpace(defaultModelId) Then
                    defaultModelId = modelId
                End If

                If StringComparer.Ordinal.Equals(modelId, normalizedPreferredModelId) Then
                    Return modelId
                End If
            Next

            If Not String.IsNullOrWhiteSpace(defaultModelId) Then
                Return defaultModelId
            End If

            If Not String.IsNullOrWhiteSpace(firstModelId) Then
                Return firstModelId
            End If

            Return normalizedPreferredModelId
        End Function

        Private Shared Function NormalizeReasoningEffort(value As String) As String
            Select Case If(value, String.Empty).Trim().ToLowerInvariant()
                Case "minimal", "low", "medium", "high", "xhigh"
                    Return If(value, String.Empty).Trim().ToLowerInvariant()
                Case Else
                    Return "medium"
            End Select
        End Function

        Private Shared Function NormalizeApprovalPolicy(value As String) As String
            Select Case If(value, String.Empty).Trim().ToLowerInvariant()
                Case "untrusted", "on-failure", "on-request", "never"
                    Return If(value, String.Empty).Trim().ToLowerInvariant()
                Case Else
                    Return "on-request"
            End Select
        End Function

        Private Shared Function NormalizeSandboxMode(value As String) As String
            Select Case If(value, String.Empty).Trim().ToLowerInvariant()
                Case "workspace-write", "read-only", "danger-full-access"
                    Return If(value, String.Empty).Trim().ToLowerInvariant()
                Case Else
                    Return "workspace-write"
            End Select
        End Function

        Private Shared Function CloneTurnComposerState(source As TurnComposerThreadState) As TurnComposerThreadState
            Return New TurnComposerThreadState() With {
                .InputText = If(source?.InputText, String.Empty),
                .ModelId = If(source?.ModelId, String.Empty),
                .ReasoningEffort = If(source?.ReasoningEffort, String.Empty),
                .ApprovalPolicy = If(source?.ApprovalPolicy, String.Empty),
                .Sandbox = If(source?.Sandbox, String.Empty)
            }
        End Function
    End Class
End Namespace
