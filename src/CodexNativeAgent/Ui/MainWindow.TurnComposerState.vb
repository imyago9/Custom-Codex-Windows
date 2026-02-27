Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Text.Json.Nodes
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
        Private ReadOnly _turnComposerPreferencesByThreadKey As New Dictionary(Of String, TurnComposerThreadPreferenceSettings)(StringComparer.Ordinal)
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
            SyncTurnComposerPreferenceForState(stateKey, _turnComposerStatesByThreadKey(stateKey))
        End Sub

        Private Sub RestoreTurnComposerStateForCurrentSelection(Optional reason As String = Nothing)
            Dim stateKey = ResolveTurnComposerStateKeyForCurrentSelection()
            If String.IsNullOrWhiteSpace(stateKey) Then
                Return
            End If

            Dim state = CloneTurnComposerState(GetOrCreateTurnComposerState(stateKey))
            ApplyTurnComposerPreferenceToState(stateKey, state)
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
            ApplyTurnComposerPreferenceToState(stateKey, state)
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
            SyncTurnComposerPreferenceForState(normalizedThreadId, promotedState)

            Dim pendingNextState = CloneTurnComposerState(pendingState)
            pendingNextState.InputText = String.Empty
            Dim normalizedPendingState = NormalizeTurnComposerStateForCurrentUi(pendingNextState)
            _turnComposerStatesByThreadKey(PendingNewThreadTranscriptTabId) = normalizedPendingState
            SyncTurnComposerPreferenceForState(PendingNewThreadTranscriptTabId, normalizedPendingState)
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
            ApplyTurnComposerPreferenceToState(normalizedStateKey, createdState)
            Dim normalizedCreatedState = NormalizeTurnComposerStateForCurrentUi(createdState)
            _turnComposerStatesByThreadKey(normalizedStateKey) = normalizedCreatedState
            Return normalizedCreatedState
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

        Private Function ExportTurnComposerThreadPreferencesForSettings() As Dictionary(Of String, TurnComposerThreadPreferenceSettings)
            Dim exported As New Dictionary(Of String, TurnComposerThreadPreferenceSettings)(StringComparer.Ordinal)

            For Each kvp In _turnComposerPreferencesByThreadKey
                Dim normalizedKey = If(kvp.Key, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalizedKey) Then
                    Continue For
                End If

                exported(normalizedKey) = CloneTurnComposerPreference(NormalizeTurnComposerPreferenceForStorage(kvp.Value))
            Next

            Return exported
        End Function

        Private Sub ImportTurnComposerThreadPreferencesFromSettings(source As IDictionary(Of String, TurnComposerThreadPreferenceSettings))
            _turnComposerPreferencesByThreadKey.Clear()

            If source Is Nothing Then
                Return
            End If

            For Each kvp In source
                Dim normalizedKey = If(kvp.Key, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalizedKey) Then
                    Continue For
                End If

                _turnComposerPreferencesByThreadKey(normalizedKey) =
                    CloneTurnComposerPreference(NormalizeTurnComposerPreferenceForStorage(kvp.Value))
            Next
        End Sub

        Private Sub SyncTurnComposerPreferenceForState(stateKey As String, state As TurnComposerThreadState)
            Dim normalizedStateKey = If(stateKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedStateKey) Then
                Return
            End If

            If state Is Nothing Then
                Return
            End If

            Dim existingPreference As TurnComposerThreadPreferenceSettings = Nothing
            _turnComposerPreferencesByThreadKey.TryGetValue(normalizedStateKey, existingPreference)

            _turnComposerPreferencesByThreadKey(normalizedStateKey) =
                NormalizeTurnComposerPreferenceForStorage(
                    New TurnComposerThreadPreferenceSettings() With {
                        .ModelId = If(state.ModelId, String.Empty),
                        .ReasoningEffort = If(state.ReasoningEffort, String.Empty),
                        .ApprovalPolicy = If(state.ApprovalPolicy, String.Empty),
                        .Sandbox = If(state.Sandbox, String.Empty),
                        .CachedContextTurnId = If(existingPreference?.CachedContextTurnId, String.Empty),
                        .CachedContextTokenUsageJson = If(existingPreference?.CachedContextTokenUsageJson, String.Empty)
                    })
        End Sub

        Private Sub ApplyTurnComposerPreferenceToState(stateKey As String, state As TurnComposerThreadState)
            Dim normalizedStateKey = If(stateKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedStateKey) OrElse state Is Nothing Then
                Return
            End If

            Dim preference As TurnComposerThreadPreferenceSettings = Nothing
            If Not _turnComposerPreferencesByThreadKey.TryGetValue(normalizedStateKey, preference) OrElse preference Is Nothing Then
                Return
            End If

            state.ModelId = If(preference.ModelId, String.Empty)
            state.ReasoningEffort = If(preference.ReasoningEffort, String.Empty)
            state.ApprovalPolicy = If(preference.ApprovalPolicy, String.Empty)
            state.Sandbox = If(preference.Sandbox, String.Empty)
        End Sub

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

        Private Shared Function NormalizeTurnComposerPreferenceForStorage(value As TurnComposerThreadPreferenceSettings) As TurnComposerThreadPreferenceSettings
            Dim source = If(value, New TurnComposerThreadPreferenceSettings())
            Return New TurnComposerThreadPreferenceSettings() With {
                .ModelId = If(source.ModelId, String.Empty).Trim(),
                .ReasoningEffort = NormalizeReasoningEffort(source.ReasoningEffort),
                .ApprovalPolicy = NormalizeApprovalPolicy(source.ApprovalPolicy),
                .Sandbox = NormalizeSandboxMode(source.Sandbox),
                .CachedContextTurnId = If(source.CachedContextTurnId, String.Empty).Trim(),
                .CachedContextTokenUsageJson = NormalizeCachedContextTokenUsageJson(source.CachedContextTokenUsageJson)
            }
        End Function

        Private Shared Function CloneTurnComposerPreference(source As TurnComposerThreadPreferenceSettings) As TurnComposerThreadPreferenceSettings
            Return New TurnComposerThreadPreferenceSettings() With {
                .ModelId = If(source?.ModelId, String.Empty),
                .ReasoningEffort = If(source?.ReasoningEffort, String.Empty),
                .ApprovalPolicy = If(source?.ApprovalPolicy, String.Empty),
                .Sandbox = If(source?.Sandbox, String.Empty),
                .CachedContextTurnId = If(source?.CachedContextTurnId, String.Empty),
                .CachedContextTokenUsageJson = If(source?.CachedContextTokenUsageJson, String.Empty)
            }
        End Function

        Private Shared Function NormalizeCachedContextTokenUsageJson(value As String) As String
            Dim normalized = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalized) Then
                Return String.Empty
            End If

            Try
                Dim parsed = JsonNode.Parse(normalized)
                Dim parsedObject = TryCast(parsed, JsonObject)
                If parsedObject Is Nothing Then
                    Return String.Empty
                End If

                Return parsedObject.ToJsonString()
            Catch
                Return String.Empty
            End Try
        End Function

        Private Sub CacheThreadContextUsageSnapshot(threadId As String, turnId As String, tokenUsage As JsonObject)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse tokenUsage Is Nothing Then
                Return
            End If

            Dim preference As TurnComposerThreadPreferenceSettings = Nothing
            If Not _turnComposerPreferencesByThreadKey.TryGetValue(normalizedThreadId, preference) OrElse preference Is Nothing Then
                preference = New TurnComposerThreadPreferenceSettings()
            Else
                preference = CloneTurnComposerPreference(preference)
            End If

            preference.CachedContextTurnId = If(turnId, String.Empty).Trim()
            preference.CachedContextTokenUsageJson = tokenUsage.ToJsonString()
            _turnComposerPreferencesByThreadKey(normalizedThreadId) = NormalizeTurnComposerPreferenceForStorage(preference)
        End Sub

        Private Function TryGetCachedThreadContextUsageSnapshot(threadId As String,
                                                                ByRef turnId As String,
                                                                ByRef tokenUsage As JsonObject) As Boolean
            turnId = String.Empty
            tokenUsage = Nothing

            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return False
            End If

            Dim preference As TurnComposerThreadPreferenceSettings = Nothing
            If Not _turnComposerPreferencesByThreadKey.TryGetValue(normalizedThreadId, preference) OrElse preference Is Nothing Then
                Return False
            End If

            Dim cachedJson = NormalizeCachedContextTokenUsageJson(preference.CachedContextTokenUsageJson)
            If String.IsNullOrWhiteSpace(cachedJson) Then
                Return False
            End If

            Try
                Dim parsed = JsonNode.Parse(cachedJson)
                Dim parsedObject = TryCast(parsed, JsonObject)
                If parsedObject Is Nothing Then
                    Return False
                End If

                turnId = If(preference.CachedContextTurnId, String.Empty).Trim()
                tokenUsage = parsedObject
                Return True
            Catch
                Return False
            End Try
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
