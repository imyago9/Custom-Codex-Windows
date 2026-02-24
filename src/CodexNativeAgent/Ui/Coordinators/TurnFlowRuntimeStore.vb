Imports System.Globalization
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.Json.Nodes

Namespace CodexNativeAgent.Ui.Coordinators
    Public NotInheritable Class ThreadRuntimeState
        Public Property ThreadId As String = String.Empty
        Public Property LatestTurnId As String = String.Empty
        Public Property IsArchived As Boolean
        Public Property StatusSummary As String = String.Empty
        Public Property LastEventMethod As String = String.Empty
        Public Property TokenUsage As JsonObject
    End Class

    Public NotInheritable Class TurnRuntimeState
        Public Property ThreadId As String = String.Empty
        Public Property TurnId As String = String.Empty
        Public Property TurnStatus As String = String.Empty
        Public Property IsCompleted As Boolean
        Public Property LastErrorMessage As String = String.Empty
        Public Property DiffSummary As String = String.Empty
        Public Property PlanSummary As String = String.Empty
        Public Property LastEventMethod As String = String.Empty
    End Class

    Public NotInheritable Class TurnItemRuntimeState
        Public Property ScopedItemKey As String = String.Empty
        Public Property ThreadId As String = String.Empty
        Public Property TurnId As String = String.Empty
        Public Property ItemId As String = String.Empty
        Public Property ItemType As String = String.Empty
        Public Property Status As String = String.Empty
        Public Property StartedAt As DateTimeOffset?
        Public Property CompletedAt As DateTimeOffset?
        Public Property RawItemPayload As JsonObject
        Public Property IsCompleted As Boolean

        Public Property AgentMessageText As String = String.Empty
        Public Property AgentMessagePhase As String = String.Empty

        Public Property PlanStreamText As String = String.Empty
        Public Property PlanFinalText As String = String.Empty

        Public Property ReasoningSummaryText As String = String.Empty
        Public ReadOnly Property ReasoningSummaryParts As New List(Of String)()
        Public Property ReasoningContentText As String = String.Empty

        Public Property CommandText As String = String.Empty
        Public Property CommandCwd As String = String.Empty
        Public Property CommandOutputText As String = String.Empty
        Public Property CommandExitCode As Integer?
        Public Property CommandDurationMs As Long?

        Public Property FileChangeOutputText As String = String.Empty
        Public Property FileChangeStatus As String = String.Empty
        Public Property FileChangeChanges As JsonArray

        Public Property GenericText As String = String.Empty

        Public Property PendingApprovalCount As Integer
        Public Property PendingApprovalSummary As String = String.Empty
        Public Property IsScopeInferred As Boolean
    End Class

    Public NotInheritable Class PendingApprovalRuntimeState
        Public Property ApprovalKey As String = String.Empty
        Public Property RequestIdKey As String = String.Empty
        Public Property RequestId As JsonNode
        Public Property MethodName As String = String.Empty
        Public Property RequestKind As ApprovalRequestKind
        Public Property ThreadId As String = String.Empty
        Public Property TurnId As String = String.Empty
        Public Property ScopedItemKey As String = String.Empty
        Public Property ItemId As String = String.Empty
        Public Property IsResolved As Boolean
        Public Property Decision As String = String.Empty
        Public Property ReceivedAt As DateTimeOffset
        Public Property ResolvedAt As DateTimeOffset?
        Public Property RawParams As JsonObject
    End Class

    Public NotInheritable Class TurnFlowReduceResult
        Public Sub New([event] As TurnFlowEvent)
            Me.EventPayload = [event]
        End Sub

        Public ReadOnly Property EventPayload As TurnFlowEvent
        Public Property UpdatedThread As ThreadRuntimeState
        Public Property UpdatedTurn As TurnRuntimeState
        Public Property UpdatedItem As TurnItemRuntimeState
        Public Property PendingApproval As PendingApprovalRuntimeState
        Public Property ResolvedApproval As PendingApprovalRuntimeState
        Public Property IsUnknownEvent As Boolean
        Public Property IsSkipped As Boolean
        Public Property SkipReason As String = String.Empty
        Public Property IsFirstSeenItem As Boolean
        Public Property DeltaAppendedLength As Integer
        Public Property DeltaPreview As String = String.Empty
        Public Property CompletionReplaced As Boolean
    End Class

    Public NotInheritable Class TurnFlowRuntimeStore
        Private ReadOnly _threadsById As New Dictionary(Of String, ThreadRuntimeState)(StringComparer.Ordinal)
        Private ReadOnly _turnsById As New Dictionary(Of String, TurnRuntimeState)(StringComparer.Ordinal)
        Private ReadOnly _itemsById As New Dictionary(Of String, TurnItemRuntimeState)(StringComparer.Ordinal)
        Private ReadOnly _itemKeyByItemId As New Dictionary(Of String, String)(StringComparer.Ordinal)
        Private ReadOnly _turnItemOrder As New Dictionary(Of String, List(Of String))(StringComparer.Ordinal)
        Private ReadOnly _pendingApprovals As New Dictionary(Of String, PendingApprovalRuntimeState)(StringComparer.Ordinal)
        Private ReadOnly _approvalKeyByRequestId As New Dictionary(Of String, String)(StringComparer.Ordinal)

        Public ReadOnly Property ThreadsById As IReadOnlyDictionary(Of String, ThreadRuntimeState)
            Get
                Return _threadsById
            End Get
        End Property

        Public ReadOnly Property TurnsById As IReadOnlyDictionary(Of String, TurnRuntimeState)
            Get
                Return _turnsById
            End Get
        End Property

        Public ReadOnly Property ItemsById As IReadOnlyDictionary(Of String, TurnItemRuntimeState)
            Get
                Return _itemsById
            End Get
        End Property

        Public ReadOnly Property TurnItemOrder As IReadOnlyDictionary(Of String, List(Of String))
            Get
                Return _turnItemOrder
            End Get
        End Property

        Public ReadOnly Property PendingApprovals As IReadOnlyDictionary(Of String, PendingApprovalRuntimeState)
            Get
                Return _pendingApprovals
            End Get
        End Property

        Public Function HasActiveTurn(threadId As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(GetActiveTurnId(threadId))
        End Function

        Public Function GetActiveTurnId(threadId As String,
                                        Optional preferredTurnId As String = Nothing) As String
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return String.Empty
            End If

            Dim normalizedPreferredTurnId = NormalizeIdentifier(preferredTurnId)
            If Not String.IsNullOrWhiteSpace(normalizedPreferredTurnId) Then
                Dim preferredTurn = GetTurnState(normalizedThreadId, normalizedPreferredTurnId)
                If preferredTurn IsNot Nothing AndAlso Not preferredTurn.IsCompleted Then
                    Return preferredTurn.TurnId
                End If
            End If

            Dim latestTurnId = GetLatestTurnId(normalizedThreadId)
            If Not String.IsNullOrWhiteSpace(latestTurnId) Then
                Dim latestTurn = GetTurnState(normalizedThreadId, latestTurnId)
                If latestTurn IsNot Nothing AndAlso Not latestTurn.IsCompleted Then
                    Return latestTurn.TurnId
                End If
            End If

            Dim activeTurnCandidates As New List(Of String)()
            For Each pair In _turnsById
                Dim turn = pair.Value
                If turn Is Nothing OrElse turn.IsCompleted Then
                    Continue For
                End If

                If Not StringComparer.Ordinal.Equals(turn.ThreadId, normalizedThreadId) Then
                    Continue For
                End If

                activeTurnCandidates.Add(turn.TurnId)
            Next

            If activeTurnCandidates.Count = 0 Then
                Return String.Empty
            End If

            activeTurnCandidates.Sort(StringComparer.Ordinal)
            Return activeTurnCandidates(0)
        End Function

        Public Function GetLatestTurnId(threadId As String) As String
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return String.Empty
            End If

            Dim thread As ThreadRuntimeState = Nothing
            If _threadsById.TryGetValue(normalizedThreadId, thread) AndAlso thread IsNot Nothing Then
                Dim latestTurnId = NormalizeIdentifier(thread.LatestTurnId)
                If Not String.IsNullOrWhiteSpace(latestTurnId) Then
                    Return latestTurnId
                End If
            End If

            Dim turnCandidates As New List(Of String)()
            For Each pair In _turnsById
                Dim turn = pair.Value
                If turn Is Nothing Then
                    Continue For
                End If

                If Not StringComparer.Ordinal.Equals(turn.ThreadId, normalizedThreadId) Then
                    Continue For
                End If

                turnCandidates.Add(turn.TurnId)
            Next

            If turnCandidates.Count = 0 Then
                Return String.Empty
            End If

            turnCandidates.Sort(StringComparer.Ordinal)
            Return turnCandidates(turnCandidates.Count - 1)
        End Function

        Public Function GetTurnState(threadId As String, turnId As String) As TurnRuntimeState
            Dim turnKey = BuildTurnKey(threadId, turnId)
            If String.IsNullOrWhiteSpace(turnKey) Then
                Return Nothing
            End If

            Dim turn As TurnRuntimeState = Nothing
            If _turnsById.TryGetValue(turnKey, turn) Then
                Return turn
            End If

            Return Nothing
        End Function

        Public Sub Reset()
            _threadsById.Clear()
            _turnsById.Clear()
            _itemsById.Clear()
            _itemKeyByItemId.Clear()
            _turnItemOrder.Clear()
            _pendingApprovals.Clear()
            _approvalKeyByRequestId.Clear()
        End Sub

        Public Function Reduce([event] As TurnFlowEvent) As TurnFlowReduceResult
            Dim result As New TurnFlowReduceResult([event])
            If [event] Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "No event payload was provided."
                Return result
            End If

            If TypeOf [event] Is ThreadStartedNotificationEvent Then
                ReduceThreadStartedNotification(DirectCast([event], ThreadStartedNotificationEvent), result)
                Return result
            End If

            If TypeOf [event] Is ThreadArchivedNotificationEvent Then
                ReduceThreadArchivedNotification(DirectCast([event], ThreadArchivedNotificationEvent), result, isArchived:=True)
                Return result
            End If

            If TypeOf [event] Is ThreadUnarchivedNotificationEvent Then
                ReduceThreadArchivedNotification(DirectCast([event], ThreadUnarchivedNotificationEvent), result, isArchived:=False)
                Return result
            End If

            If TypeOf [event] Is ThreadStatusChangedNotificationEvent Then
                ReduceThreadStatusChangedNotification(DirectCast([event], ThreadStatusChangedNotificationEvent), result)
                Return result
            End If

            If TypeOf [event] Is TurnStartedEvent Then
                ReduceTurnStarted(DirectCast([event], TurnStartedEvent), result)
                Return result
            End If

            If TypeOf [event] Is TurnCompletedEvent Then
                ReduceTurnCompleted(DirectCast([event], TurnCompletedEvent), result)
                Return result
            End If

            If TypeOf [event] Is TurnDiffUpdatedEvent Then
                ReduceTurnDiffUpdated(DirectCast([event], TurnDiffUpdatedEvent), result)
                Return result
            End If

            If TypeOf [event] Is TurnPlanUpdatedEvent Then
                ReduceTurnPlanUpdated(DirectCast([event], TurnPlanUpdatedEvent), result)
                Return result
            End If

            If TypeOf [event] Is ThreadTokenUsageUpdatedEvent Then
                ReduceThreadTokenUsageUpdated(DirectCast([event], ThreadTokenUsageUpdatedEvent), result)
                Return result
            End If

            If TypeOf [event] Is ItemLifecycleEvent Then
                ReduceItemLifecycle(DirectCast([event], ItemLifecycleEvent), result)
                Return result
            End If

            If TypeOf [event] Is ItemDeltaEvent Then
                ReduceItemDelta(DirectCast([event], ItemDeltaEvent), result)
                Return result
            End If

            If TypeOf [event] Is ApprovalRequestEvent Then
                ReduceApprovalRequest(DirectCast([event], ApprovalRequestEvent), result)
                Return result
            End If

            If TypeOf [event] Is ErrorEvent Then
                ReduceErrorEvent(DirectCast([event], ErrorEvent), result)
                Return result
            End If

            If TypeOf [event] Is GenericServerNotificationEvent OrElse
               TypeOf [event] Is GenericServerRequestEvent OrElse
               TypeOf [event] Is AccountLoginCompletedNotificationEvent OrElse
               TypeOf [event] Is AccountUpdatedNotificationEvent OrElse
               TypeOf [event] Is AccountRateLimitsUpdatedNotificationEvent OrElse
               TypeOf [event] Is ModelReroutedNotificationEvent OrElse
               TypeOf [event] Is AppListUpdatedNotificationEvent OrElse
               TypeOf [event] Is McpOauthLoginCompletedNotificationEvent OrElse
               TypeOf [event] Is WindowsSandboxSetupCompletedNotificationEvent OrElse
               TypeOf [event] Is FuzzyFileSearchSessionUpdatedNotificationEvent OrElse
               TypeOf [event] Is FuzzyFileSearchSessionCompletedNotificationEvent Then
                Return result
            End If

            result.IsSkipped = True
            result.SkipReason = $"Unhandled event payload type '{[event].GetType().Name}'."
            Return result
        End Function

        Public Function ResolveApprovalDecision(requestId As JsonNode, decision As String) As TurnFlowReduceResult
            Dim result As New TurnFlowReduceResult(Nothing)
            Dim requestIdKey = BuildRequestIdKey(requestId)
            If String.IsNullOrWhiteSpace(requestIdKey) Then
                result.IsSkipped = True
                result.SkipReason = "Approval resolution skipped: missing request id."
                Return result
            End If

            Dim approvalKey As String = Nothing
            If Not _approvalKeyByRequestId.TryGetValue(requestIdKey, approvalKey) OrElse
               String.IsNullOrWhiteSpace(approvalKey) Then
                result.IsSkipped = True
                result.SkipReason = $"Approval resolution skipped: unknown request id '{requestIdKey}'."
                Return result
            End If

            Dim approval As PendingApprovalRuntimeState = Nothing
            If Not _pendingApprovals.TryGetValue(approvalKey, approval) OrElse approval Is Nothing Then
                _approvalKeyByRequestId.Remove(requestIdKey)
                result.IsSkipped = True
                result.SkipReason = $"Approval resolution skipped: missing approval key '{approvalKey}'."
                Return result
            End If

            approval.IsResolved = True
            approval.Decision = If(decision, String.Empty).Trim()
            approval.ResolvedAt = DateTimeOffset.UtcNow

            result.ResolvedApproval = approval

            _pendingApprovals.Remove(approvalKey)
            _approvalKeyByRequestId.Remove(requestIdKey)

            Dim scopedItemKey = NormalizeIdentifier(approval.ScopedItemKey)
            If String.IsNullOrWhiteSpace(scopedItemKey) Then
                _itemKeyByItemId.TryGetValue(NormalizeIdentifier(approval.ItemId), scopedItemKey)
            End If

            Dim item As TurnItemRuntimeState = Nothing
            If Not String.IsNullOrWhiteSpace(scopedItemKey) AndAlso
               _itemsById.TryGetValue(scopedItemKey, item) AndAlso item IsNot Nothing Then
                RefreshItemPendingApprovalState(item)
                result.UpdatedItem = item
            End If

            Return result
        End Function

        Private Sub ReduceThreadStartedNotification(evt As ThreadStartedNotificationEvent, result As TurnFlowReduceResult)
            If evt Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "Thread started notification was null."
                Return
            End If

            Dim normalizedThreadId = NormalizeIdentifier(evt.ThreadId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) AndAlso evt.ThreadObject IsNot Nothing Then
                normalizedThreadId = ReadStringFirst(evt.ThreadObject, "id")
            End If

            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                result.IsSkipped = True
                result.SkipReason = "Thread started notification is missing thread id."
                Return
            End If

            Dim thread = EnsureThread(normalizedThreadId)
            thread.IsArchived = False
            thread.LastEventMethod = evt.MethodName
            result.UpdatedThread = thread
        End Sub

        Private Sub ReduceThreadArchivedNotification(evt As TurnFlowEvent,
                                                     result As TurnFlowReduceResult,
                                                     isArchived As Boolean)
            If evt Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "Thread archive notification was null."
                Return
            End If

            Dim threadId As String = String.Empty
            If TypeOf evt Is ThreadArchivedNotificationEvent Then
                threadId = DirectCast(evt, ThreadArchivedNotificationEvent).ThreadId
            ElseIf TypeOf evt Is ThreadUnarchivedNotificationEvent Then
                threadId = DirectCast(evt, ThreadUnarchivedNotificationEvent).ThreadId
            End If

            threadId = NormalizeIdentifier(threadId)
            If String.IsNullOrWhiteSpace(threadId) Then
                result.IsSkipped = True
                result.SkipReason = "Thread archive notification is missing thread id."
                Return
            End If

            Dim thread = EnsureThread(threadId)
            thread.IsArchived = isArchived
            thread.LastEventMethod = evt.MethodName
            result.UpdatedThread = thread
        End Sub

        Private Sub ReduceThreadStatusChangedNotification(evt As ThreadStatusChangedNotificationEvent,
                                                          result As TurnFlowReduceResult)
            If evt Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "Thread status changed notification was null."
                Return
            End If

            Dim threadId = NormalizeIdentifier(evt.ThreadId)
            If String.IsNullOrWhiteSpace(threadId) Then
                result.IsSkipped = True
                result.SkipReason = "Thread status changed notification is missing thread id."
                Return
            End If

            Dim thread = EnsureThread(threadId)
            thread.StatusSummary = If(evt.StatusSummary, String.Empty)
            thread.LastEventMethod = evt.MethodName
            result.UpdatedThread = thread
        End Sub

        Private Sub ReduceTurnStarted(evt As TurnStartedEvent, result As TurnFlowReduceResult)
            If evt Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "Turn started event was null."
                Return
            End If

            Dim resolvedThreadId As String = String.Empty
            Dim resolvedTurnId As String = String.Empty
            If Not TryResolveTurnScope(evt.ThreadId, evt.TurnId, resolvedThreadId, resolvedTurnId) Then
                result.IsSkipped = True
                result.SkipReason = "Turn started event scope could not be resolved."
                Return
            End If

            Dim turn = EnsureTurn(resolvedThreadId, resolvedTurnId)
            turn.IsCompleted = False
            turn.TurnStatus = "in_progress"
            turn.LastErrorMessage = String.Empty
            turn.LastEventMethod = evt.MethodName

            result.UpdatedTurn = turn
            result.UpdatedThread = EnsureThread(resolvedThreadId)
        End Sub

        Private Sub ReduceTurnCompleted(evt As TurnCompletedEvent, result As TurnFlowReduceResult)
            If evt Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "Turn completed event was null."
                Return
            End If

            Dim resolvedThreadId As String = String.Empty
            Dim resolvedTurnId As String = String.Empty
            If Not TryResolveTurnScope(evt.ThreadId, evt.TurnId, resolvedThreadId, resolvedTurnId) Then
                result.IsSkipped = True
                result.SkipReason = "Turn completed event scope could not be resolved."
                Return
            End If

            Dim turn = EnsureTurn(resolvedThreadId, resolvedTurnId)
            turn.IsCompleted = True
            turn.TurnStatus = NormalizeStatus(evt.Status, "completed")
            If Not String.IsNullOrWhiteSpace(evt.ErrorMessage) Then
                turn.LastErrorMessage = evt.ErrorMessage
            End If
            turn.LastEventMethod = evt.MethodName

            ClearPendingApprovalsForTurn(resolvedThreadId, resolvedTurnId)

            result.UpdatedTurn = turn
            result.UpdatedThread = EnsureThread(resolvedThreadId)
        End Sub

        Private Sub ReduceTurnDiffUpdated(evt As TurnDiffUpdatedEvent, result As TurnFlowReduceResult)
            If evt Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "Turn diff update event was null."
                Return
            End If

            Dim resolvedThreadId As String = String.Empty
            Dim resolvedTurnId As String = String.Empty
            If Not TryResolveTurnScope(evt.ThreadId, evt.TurnId, resolvedThreadId, resolvedTurnId) Then
                result.IsSkipped = True
                result.SkipReason = "Turn diff update scope could not be resolved."
                Return
            End If

            Dim turn = EnsureTurn(resolvedThreadId, resolvedTurnId)
            turn.DiffSummary = If(evt.SummaryText, String.Empty)
            turn.LastEventMethod = evt.MethodName
            result.UpdatedTurn = turn
            result.UpdatedThread = EnsureThread(resolvedThreadId)
        End Sub

        Private Sub ReduceTurnPlanUpdated(evt As TurnPlanUpdatedEvent, result As TurnFlowReduceResult)
            If evt Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "Turn plan update event was null."
                Return
            End If

            Dim resolvedThreadId As String = String.Empty
            Dim resolvedTurnId As String = String.Empty
            If Not TryResolveTurnScope(evt.ThreadId, evt.TurnId, resolvedThreadId, resolvedTurnId) Then
                result.IsSkipped = True
                result.SkipReason = "Turn plan update scope could not be resolved."
                Return
            End If

            Dim turn = EnsureTurn(resolvedThreadId, resolvedTurnId)
            turn.PlanSummary = If(evt.SummaryText, String.Empty)
            turn.LastEventMethod = evt.MethodName
            result.UpdatedTurn = turn
            result.UpdatedThread = EnsureThread(resolvedThreadId)
        End Sub

        Private Sub ReduceThreadTokenUsageUpdated(evt As ThreadTokenUsageUpdatedEvent, result As TurnFlowReduceResult)
            If evt Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "Token usage update event was null."
                Return
            End If

            Dim resolvedThreadId = NormalizeIdentifier(evt.ThreadId)
            Dim resolvedTurnId = NormalizeIdentifier(evt.TurnId)

            If String.IsNullOrWhiteSpace(resolvedThreadId) AndAlso
               Not String.IsNullOrWhiteSpace(resolvedTurnId) Then
                Dim inferredThreadId As String = String.Empty
                Dim inferredTurnId As String = String.Empty
                If TryResolveTurnScope(String.Empty, resolvedTurnId, inferredThreadId, inferredTurnId) Then
                    resolvedThreadId = inferredThreadId
                End If
            End If

            If String.IsNullOrWhiteSpace(resolvedThreadId) Then
                result.IsSkipped = True
                result.SkipReason = "Token usage update is missing threadId."
                Return
            End If

            Dim thread = EnsureThread(resolvedThreadId)
            thread.TokenUsage = CloneJsonObject(evt.TokenUsage)
            thread.LastEventMethod = evt.MethodName
            If Not String.IsNullOrWhiteSpace(resolvedTurnId) Then
                thread.LatestTurnId = resolvedTurnId
            End If

            If Not String.IsNullOrWhiteSpace(resolvedTurnId) Then
                Dim turn = EnsureTurn(resolvedThreadId, resolvedTurnId)
                turn.LastEventMethod = evt.MethodName
                result.UpdatedTurn = turn
            End If

            result.UpdatedThread = thread
        End Sub

        Private Sub ReduceItemLifecycle(evt As ItemLifecycleEvent, result As TurnFlowReduceResult)
            If evt Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "Item lifecycle event was null."
                Return
            End If

            Dim itemId = NormalizeIdentifier(evt.ItemId)
            If String.IsNullOrWhiteSpace(itemId) Then
                result.IsSkipped = True
                result.SkipReason = $"{evt.MethodName} ignored: missing item.id."
                Return
            End If

            Dim resolvedThreadId As String = Nothing
            Dim resolvedTurnId As String = Nothing
            If Not TryResolveItemScope(itemId,
                                       evt.ThreadId,
                                       evt.TurnId,
                                       resolvedThreadId,
                                       resolvedTurnId,
                                       result) Then
                Return
            End If

            Dim payloadType = ReadStringFirst(evt.ItemObject, "type")
            Dim item = EnsureItem(itemId,
                                  resolvedThreadId,
                                  resolvedTurnId,
                                  payloadType,
                                  result)
            If item Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = $"{evt.MethodName} ignored: unresolved item scope."
                Return
            End If

            item.ItemType = NormalizeItemType(If(String.IsNullOrWhiteSpace(payloadType), item.ItemType, payloadType))
            item.RawItemPayload = CloneJsonObject(evt.ItemObject)

            If evt.IsCompleted Then
                item.IsCompleted = True
                item.CompletedAt = DateTimeOffset.UtcNow
                item.Status = NormalizeStatus(ReadStringFirst(evt.ItemObject, "status"), "completed")
            Else
                item.IsCompleted = False
                If Not item.StartedAt.HasValue Then
                    item.StartedAt = DateTimeOffset.UtcNow
                End If
                item.Status = NormalizeStatus(ReadStringFirst(evt.ItemObject, "status"), "in_progress")
            End If

            Dim completionReplaced = ApplyItemPayload(item, evt.ItemObject, evt.IsCompleted)
            If evt.IsCompleted Then
                result.CompletionReplaced = completionReplaced
            End If

            EnsureItemOrder(item)

            result.UpdatedItem = item
            result.UpdatedTurn = EnsureTurn(item.ThreadId, item.TurnId)
            result.UpdatedThread = EnsureThread(item.ThreadId)
        End Sub

        Private Sub ReduceItemDelta(evt As ItemDeltaEvent, result As TurnFlowReduceResult)
            If evt Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "Item delta event was null."
                Return
            End If

            Dim itemId = NormalizeIdentifier(evt.ItemId)
            If String.IsNullOrWhiteSpace(itemId) Then
                result.IsSkipped = True
                result.SkipReason = $"{evt.MethodName} ignored: missing itemId."
                Return
            End If

            Dim resolvedThreadId As String = Nothing
            Dim resolvedTurnId As String = Nothing
            If Not TryResolveItemScope(itemId,
                                       evt.ThreadId,
                                       evt.TurnId,
                                       resolvedThreadId,
                                       resolvedTurnId,
                                       result) Then
                Return
            End If

            Dim item = EnsureItem(itemId,
                                  resolvedThreadId,
                                  resolvedTurnId,
                                  ItemTypeFromDeltaKind(evt.DeltaKind),
                                  result)
            If item Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = $"{evt.MethodName} ignored: unresolved item scope."
                Return
            End If

            Dim deltaText = If(evt.DeltaText, String.Empty)
            Dim summaryPartText = If(evt.SummaryPartText, String.Empty)

            Select Case evt.DeltaKind
                Case ItemDeltaKind.AgentMessage
                    item.ItemType = NormalizeItemType(If(String.IsNullOrWhiteSpace(item.ItemType), "agentMessage", item.ItemType))
                    item.AgentMessageText &= deltaText
                    result.DeltaPreview = PreviewDelta(deltaText)
                    result.DeltaAppendedLength = deltaText.Length

                Case ItemDeltaKind.Plan
                    item.ItemType = NormalizeItemType(If(String.IsNullOrWhiteSpace(item.ItemType), "plan", item.ItemType))
                    item.PlanStreamText &= deltaText
                    result.DeltaPreview = PreviewDelta(deltaText)
                    result.DeltaAppendedLength = deltaText.Length

                Case ItemDeltaKind.ReasoningSummaryText
                    item.ItemType = NormalizeItemType(If(String.IsNullOrWhiteSpace(item.ItemType), "reasoning", item.ItemType))
                    item.ReasoningSummaryText &= deltaText
                    result.DeltaPreview = PreviewDelta(deltaText)
                    result.DeltaAppendedLength = deltaText.Length

                Case ItemDeltaKind.ReasoningSummaryPartAdded
                    item.ItemType = NormalizeItemType(If(String.IsNullOrWhiteSpace(item.ItemType), "reasoning", item.ItemType))
                    If Not String.IsNullOrWhiteSpace(summaryPartText) Then
                        item.ReasoningSummaryParts.Add(summaryPartText)
                        If String.IsNullOrWhiteSpace(item.ReasoningSummaryText) Then
                            item.ReasoningSummaryText = summaryPartText
                        Else
                            item.ReasoningSummaryText &= Environment.NewLine & summaryPartText
                        End If
                        result.DeltaPreview = PreviewDelta(summaryPartText)
                        result.DeltaAppendedLength = summaryPartText.Length
                    ElseIf Not String.IsNullOrWhiteSpace(item.ReasoningSummaryText) AndAlso
                        Not item.ReasoningSummaryText.EndsWith(Environment.NewLine, StringComparison.Ordinal) Then
                        ' summaryPartAdded can be a boundary marker without inline text.
                        item.ReasoningSummaryText &= Environment.NewLine
                    End If

                Case ItemDeltaKind.ReasoningText
                    item.ItemType = NormalizeItemType(If(String.IsNullOrWhiteSpace(item.ItemType), "reasoning", item.ItemType))
                    item.ReasoningContentText &= deltaText
                    result.DeltaPreview = PreviewDelta(deltaText)
                    result.DeltaAppendedLength = deltaText.Length

                Case ItemDeltaKind.CommandExecutionOutput
                    item.ItemType = NormalizeItemType(If(String.IsNullOrWhiteSpace(item.ItemType), "commandExecution", item.ItemType))
                    item.CommandOutputText &= deltaText
                    result.DeltaPreview = PreviewDelta(deltaText)
                    result.DeltaAppendedLength = deltaText.Length

                Case ItemDeltaKind.FileChangeOutput
                    item.ItemType = NormalizeItemType(If(String.IsNullOrWhiteSpace(item.ItemType), "fileChange", item.ItemType))
                    item.FileChangeOutputText &= deltaText
                    result.DeltaPreview = PreviewDelta(deltaText)
                    result.DeltaAppendedLength = deltaText.Length
            End Select

            item.Status = "in_progress"
            item.IsCompleted = False
            If Not item.StartedAt.HasValue Then
                item.StartedAt = DateTimeOffset.UtcNow
            End If

            EnsureItemOrder(item)

            result.UpdatedItem = item
            result.UpdatedTurn = EnsureTurn(item.ThreadId, item.TurnId)
            result.UpdatedThread = EnsureThread(item.ThreadId)
        End Sub

        Private Sub ReduceApprovalRequest(evt As ApprovalRequestEvent, result As TurnFlowReduceResult)
            If evt Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "Approval request event was null."
                Return
            End If

            Dim itemId = NormalizeIdentifier(evt.ItemId)
            If String.IsNullOrWhiteSpace(itemId) Then
                result.IsSkipped = True
                result.SkipReason = $"{evt.MethodName} ignored: missing itemId."
                Return
            End If

            Dim resolvedThreadId As String = Nothing
            Dim resolvedTurnId As String = Nothing
            If Not TryResolveItemScope(itemId,
                                       evt.ThreadId,
                                       evt.TurnId,
                                       resolvedThreadId,
                                       resolvedTurnId,
                                       result) Then
                Return
            End If

            Dim item = EnsureItem(itemId,
                                  resolvedThreadId,
                                  resolvedTurnId,
                                  If(evt.RequestKind = ApprovalRequestKind.CommandExecution, "commandExecution", "fileChange"),
                                  result)
            If item Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = $"{evt.MethodName} ignored: unresolved item scope."
                Return
            End If
            EnsureItemOrder(item)

            Dim requestIdKey = BuildRequestIdKey(evt.RequestId)
            If String.IsNullOrWhiteSpace(requestIdKey) Then
                result.IsSkipped = True
                result.SkipReason = $"{evt.MethodName} ignored: request id is missing."
                Return
            End If

            Dim approvalKey = BuildApprovalKey(item.ScopedItemKey, requestIdKey)
            Dim approval As New PendingApprovalRuntimeState() With {
                .ApprovalKey = approvalKey,
                .RequestIdKey = requestIdKey,
                .RequestId = CloneJsonNode(evt.RequestId),
                .MethodName = evt.MethodName,
                .RequestKind = evt.RequestKind,
                .ThreadId = resolvedThreadId,
                .TurnId = resolvedTurnId,
                .ScopedItemKey = item.ScopedItemKey,
                .ItemId = itemId,
                .IsResolved = False,
                .Decision = String.Empty,
                .ReceivedAt = DateTimeOffset.UtcNow,
                .RawParams = CloneJsonObject(evt.RawParams)
            }

            _pendingApprovals(approvalKey) = approval
            _approvalKeyByRequestId(requestIdKey) = approvalKey

            RefreshItemPendingApprovalState(item)

            result.PendingApproval = approval
            result.UpdatedItem = item
            result.UpdatedTurn = EnsureTurn(item.ThreadId, item.TurnId)
            result.UpdatedThread = EnsureThread(item.ThreadId)
        End Sub

        Private Sub ReduceErrorEvent(evt As ErrorEvent, result As TurnFlowReduceResult)
            If evt Is Nothing Then
                result.IsSkipped = True
                result.SkipReason = "Error event was null."
                Return
            End If

            If Not String.IsNullOrWhiteSpace(evt.ThreadId) Then
                result.UpdatedThread = EnsureThread(evt.ThreadId)
            End If

            If Not String.IsNullOrWhiteSpace(evt.ThreadId) AndAlso
               Not String.IsNullOrWhiteSpace(evt.TurnId) Then
                Dim turn = EnsureTurn(evt.ThreadId, evt.TurnId)
                If Not String.IsNullOrWhiteSpace(evt.Message) Then
                    turn.LastErrorMessage = evt.Message
                End If
                If Not turn.IsCompleted Then
                    turn.TurnStatus = "failed"
                End If
                turn.LastEventMethod = evt.MethodName

                result.UpdatedTurn = turn
            End If
        End Sub

        Private Function EnsureThread(threadId As String) As ThreadRuntimeState
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return Nothing
            End If

            Dim thread As ThreadRuntimeState = Nothing
            If _threadsById.TryGetValue(normalizedThreadId, thread) AndAlso thread IsNot Nothing Then
                Return thread
            End If

            thread = New ThreadRuntimeState() With {
                .ThreadId = normalizedThreadId
            }
            _threadsById(normalizedThreadId) = thread
            Return thread
        End Function

        Private Function EnsureTurn(threadId As String, turnId As String) As TurnRuntimeState
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            Dim normalizedTurnId = NormalizeIdentifier(turnId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return Nothing
            End If

            Dim thread = EnsureThread(normalizedThreadId)
            If thread IsNot Nothing Then
                thread.LatestTurnId = normalizedTurnId
            End If

            Dim turnKey = BuildTurnKey(normalizedThreadId, normalizedTurnId)
            Dim turn As TurnRuntimeState = Nothing
            If _turnsById.TryGetValue(turnKey, turn) AndAlso turn IsNot Nothing Then
                Return turn
            End If

            turn = New TurnRuntimeState() With {
                .ThreadId = normalizedThreadId,
                .TurnId = normalizedTurnId,
                .TurnStatus = "in_progress",
                .IsCompleted = False
            }
            _turnsById(turnKey) = turn
            Return turn
        End Function

        Private Function TryResolveTurnScope(eventThreadId As String,
                                             eventTurnId As String,
                                             ByRef resolvedThreadId As String,
                                             ByRef resolvedTurnId As String) As Boolean
            resolvedThreadId = NormalizeIdentifier(eventThreadId)
            resolvedTurnId = NormalizeIdentifier(eventTurnId)

            If Not String.IsNullOrWhiteSpace(resolvedThreadId) AndAlso
               Not String.IsNullOrWhiteSpace(resolvedTurnId) Then
                Return True
            End If

            If Not String.IsNullOrWhiteSpace(resolvedThreadId) AndAlso
               String.IsNullOrWhiteSpace(resolvedTurnId) Then
                Dim thread As ThreadRuntimeState = Nothing
                If _threadsById.TryGetValue(resolvedThreadId, thread) AndAlso thread IsNot Nothing Then
                    resolvedTurnId = NormalizeIdentifier(thread.LatestTurnId)
                    If Not String.IsNullOrWhiteSpace(resolvedTurnId) Then
                        Return True
                    End If
                End If
            End If

            If String.IsNullOrWhiteSpace(resolvedThreadId) AndAlso
               Not String.IsNullOrWhiteSpace(resolvedTurnId) Then
                Dim matchedThreadId As String = String.Empty
                For Each turnPair In _turnsById
                    Dim turn = turnPair.Value
                    If turn Is Nothing Then
                        Continue For
                    End If

                    If Not StringComparer.Ordinal.Equals(turn.TurnId, resolvedTurnId) Then
                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(matchedThreadId) Then
                        matchedThreadId = turn.ThreadId
                    ElseIf Not StringComparer.Ordinal.Equals(matchedThreadId, turn.ThreadId) Then
                        matchedThreadId = String.Empty
                        Exit For
                    End If
                Next

                If Not String.IsNullOrWhiteSpace(matchedThreadId) Then
                    resolvedThreadId = matchedThreadId
                    Return True
                End If
            End If

            If String.IsNullOrWhiteSpace(resolvedThreadId) Then
                If _threadsById.Count = 1 Then
                    For Each pair In _threadsById
                        resolvedThreadId = pair.Key
                    Next
                Else
                    resolvedThreadId = "__inferred_thread__"
                End If
            End If

            If String.IsNullOrWhiteSpace(resolvedTurnId) Then
                Dim thread As ThreadRuntimeState = Nothing
                If _threadsById.TryGetValue(resolvedThreadId, thread) AndAlso thread IsNot Nothing Then
                    resolvedTurnId = NormalizeIdentifier(thread.LatestTurnId)
                End If
            End If

            If String.IsNullOrWhiteSpace(resolvedTurnId) Then
                resolvedTurnId = "__inferred_turn__"
            End If

            Return True
        End Function

        Private Function EnsureItem(itemId As String,
                                    threadId As String,
                                    turnId As String,
                                    defaultType As String,
                                    result As TurnFlowReduceResult) As TurnItemRuntimeState
            Dim normalizedItemId = NormalizeIdentifier(itemId)
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            Dim normalizedTurnId = NormalizeIdentifier(turnId)
            If String.IsNullOrWhiteSpace(normalizedItemId) OrElse
               String.IsNullOrWhiteSpace(normalizedThreadId) OrElse
               String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return Nothing
            End If

            Dim scopedItemKey = BuildItemKey(normalizedThreadId, normalizedTurnId, normalizedItemId)
            Dim item As TurnItemRuntimeState = Nothing
            If _itemsById.TryGetValue(scopedItemKey, item) AndAlso item IsNot Nothing Then
                Dim existingMappedKey As String = Nothing
                _itemKeyByItemId.TryGetValue(normalizedItemId, existingMappedKey)
                If Not StringComparer.Ordinal.Equals(existingMappedKey, scopedItemKey) Then
                    _itemKeyByItemId(normalizedItemId) = scopedItemKey
                End If
                Return item
            End If

            item = New TurnItemRuntimeState() With {
                .ScopedItemKey = scopedItemKey,
                .ItemId = normalizedItemId,
                .ThreadId = normalizedThreadId,
                .TurnId = normalizedTurnId,
                .ItemType = NormalizeItemType(defaultType),
                .Status = "in_progress",
                .StartedAt = DateTimeOffset.UtcNow,
                .IsCompleted = False,
                .IsScopeInferred = normalizedThreadId.StartsWith("__inferred_", StringComparison.Ordinal) OrElse
                                   normalizedTurnId.StartsWith("__inferred_", StringComparison.Ordinal)
            }

            _itemsById(scopedItemKey) = item
            _itemKeyByItemId(normalizedItemId) = scopedItemKey
            result.IsFirstSeenItem = True
            Return item
        End Function

        Private Sub EnsureItemOrder(item As TurnItemRuntimeState)
            If item Is Nothing Then
                Return
            End If

            Dim turnKey = BuildTurnKey(item.ThreadId, item.TurnId)
            If String.IsNullOrWhiteSpace(turnKey) Then
                Return
            End If

            Dim order As List(Of String) = Nothing
            If Not _turnItemOrder.TryGetValue(turnKey, order) OrElse order Is Nothing Then
                order = New List(Of String)()
                _turnItemOrder(turnKey) = order
            End If

            Dim alreadyPresent = False
            For Each existingKey In order
                If StringComparer.Ordinal.Equals(existingKey, item.ScopedItemKey) Then
                    alreadyPresent = True
                    Exit For
                End If
            Next

            If Not alreadyPresent Then
                order.Add(item.ScopedItemKey)
            End If
        End Sub

        Private Function TryResolveItemScope(itemId As String,
                                             eventThreadId As String,
                                             eventTurnId As String,
                                             ByRef resolvedThreadId As String,
                                             ByRef resolvedTurnId As String,
                                             result As TurnFlowReduceResult) As Boolean
            resolvedThreadId = String.Empty
            resolvedTurnId = String.Empty

            Dim normalizedItemId = NormalizeIdentifier(itemId)
            Dim normalizedThreadId = NormalizeIdentifier(eventThreadId)
            Dim normalizedTurnId = NormalizeIdentifier(eventTurnId)

            Dim existingKey As String = Nothing
            Dim existing As TurnItemRuntimeState = Nothing
            If Not String.IsNullOrWhiteSpace(normalizedItemId) AndAlso
               _itemKeyByItemId.TryGetValue(normalizedItemId, existingKey) AndAlso
               Not String.IsNullOrWhiteSpace(existingKey) Then
                _itemsById.TryGetValue(existingKey, existing)
            End If

            If existing IsNot Nothing Then
                If Not String.IsNullOrWhiteSpace(normalizedThreadId) AndAlso
                   Not String.IsNullOrWhiteSpace(normalizedTurnId) AndAlso
                   (Not StringComparer.Ordinal.Equals(normalizedThreadId, existing.ThreadId) OrElse
                    Not StringComparer.Ordinal.Equals(normalizedTurnId, existing.TurnId)) Then
                    resolvedThreadId = normalizedThreadId
                    resolvedTurnId = normalizedTurnId
                    Return True
                End If

                If Not String.IsNullOrWhiteSpace(normalizedThreadId) AndAlso
                   Not StringComparer.Ordinal.Equals(normalizedThreadId, existing.ThreadId) Then
                    normalizedThreadId = existing.ThreadId
                End If

                If Not String.IsNullOrWhiteSpace(normalizedTurnId) AndAlso
                   Not StringComparer.Ordinal.Equals(normalizedTurnId, existing.TurnId) Then
                    normalizedTurnId = existing.TurnId
                End If

                resolvedThreadId = If(String.IsNullOrWhiteSpace(normalizedThreadId), existing.ThreadId, normalizedThreadId)
                resolvedTurnId = If(String.IsNullOrWhiteSpace(normalizedTurnId), existing.TurnId, normalizedTurnId)
                Return True
            End If

            If TryInferScope(normalizedThreadId, normalizedTurnId, resolvedThreadId, resolvedTurnId) Then
                Return True
            End If

            resolvedThreadId = If(String.IsNullOrWhiteSpace(normalizedThreadId), "__inferred_thread__", normalizedThreadId)
            resolvedTurnId = If(String.IsNullOrWhiteSpace(normalizedTurnId), "__inferred_turn__", normalizedTurnId)
            result.SkipReason = $"Scope inferred for item '{normalizedItemId}' as {resolvedThreadId}:{resolvedTurnId}."
            Return True
        End Function

        Private Function TryInferScope(candidateThreadId As String,
                                       candidateTurnId As String,
                                       ByRef resolvedThreadId As String,
                                       ByRef resolvedTurnId As String) As Boolean
            resolvedThreadId = NormalizeIdentifier(candidateThreadId)
            resolvedTurnId = NormalizeIdentifier(candidateTurnId)

            If Not String.IsNullOrWhiteSpace(resolvedThreadId) AndAlso
               Not String.IsNullOrWhiteSpace(resolvedTurnId) Then
                Return True
            End If

            If Not String.IsNullOrWhiteSpace(resolvedThreadId) AndAlso String.IsNullOrWhiteSpace(resolvedTurnId) Then
                Dim thread As ThreadRuntimeState = Nothing
                If _threadsById.TryGetValue(resolvedThreadId, thread) AndAlso thread IsNot Nothing AndAlso
                   Not String.IsNullOrWhiteSpace(thread.LatestTurnId) Then
                    resolvedTurnId = thread.LatestTurnId
                    Return True
                End If
            End If

            If String.IsNullOrWhiteSpace(resolvedThreadId) AndAlso Not String.IsNullOrWhiteSpace(resolvedTurnId) Then
                Dim matchedThreadId As String = String.Empty
                For Each turnPair In _turnsById
                    Dim turn = turnPair.Value
                    If turn Is Nothing Then
                        Continue For
                    End If

                    If Not StringComparer.Ordinal.Equals(turn.TurnId, resolvedTurnId) Then
                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(matchedThreadId) Then
                        matchedThreadId = turn.ThreadId
                    ElseIf Not StringComparer.Ordinal.Equals(matchedThreadId, turn.ThreadId) Then
                        matchedThreadId = String.Empty
                        Exit For
                    End If
                Next

                If Not String.IsNullOrWhiteSpace(matchedThreadId) Then
                    resolvedThreadId = matchedThreadId
                    Return True
                End If
            End If

            If String.IsNullOrWhiteSpace(resolvedThreadId) AndAlso String.IsNullOrWhiteSpace(resolvedTurnId) Then
                Dim activeThreadId As String = String.Empty
                Dim activeTurnId As String = String.Empty

                For Each turnPair In _turnsById
                    Dim turn = turnPair.Value
                    If turn Is Nothing OrElse turn.IsCompleted Then
                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(activeTurnId) Then
                        activeThreadId = turn.ThreadId
                        activeTurnId = turn.TurnId
                    Else
                        activeThreadId = String.Empty
                        activeTurnId = String.Empty
                        Exit For
                    End If
                Next

                If Not String.IsNullOrWhiteSpace(activeThreadId) AndAlso
                   Not String.IsNullOrWhiteSpace(activeTurnId) Then
                    resolvedThreadId = activeThreadId
                    resolvedTurnId = activeTurnId
                    Return True
                End If
            End If

            Return Not String.IsNullOrWhiteSpace(resolvedThreadId) AndAlso
                   Not String.IsNullOrWhiteSpace(resolvedTurnId)
        End Function

        Private Function ApplyItemPayload(item As TurnItemRuntimeState,
                                          itemObject As JsonObject,
                                          isFinal As Boolean) As Boolean
            If item Is Nothing OrElse itemObject Is Nothing Then
                Return False
            End If

            Dim completionReplaced = False
            Dim itemType = NormalizeItemType(ReadStringFirst(itemObject, "type"))
            If Not String.IsNullOrWhiteSpace(itemType) Then
                item.ItemType = itemType
            End If

            Select Case item.ItemType
                Case "userMessage"
                    item.GenericText = ReadUserMessageText(itemObject)

                Case "agentMessage"
                    Dim phase = NormalizeIdentifier(ReadAgentMessagePhase(itemObject))
                    If Not String.IsNullOrWhiteSpace(phase) Then
                        item.AgentMessagePhase = phase
                    End If

                    Dim agentText = ReadAgentMessageText(itemObject)
                    If isFinal Then
                        If Not String.IsNullOrWhiteSpace(agentText) Then
                            completionReplaced = Not StringComparer.Ordinal.Equals(item.AgentMessageText, agentText)
                            item.AgentMessageText = agentText
                        End If
                    ElseIf Not String.IsNullOrWhiteSpace(agentText) Then
                        item.AgentMessageText = agentText
                    End If

                Case "plan"
                    Dim planText = ReadStringFirst(itemObject, "text")
                    If isFinal Then
                        If Not String.IsNullOrWhiteSpace(planText) Then
                            completionReplaced = Not StringComparer.Ordinal.Equals(item.PlanStreamText, planText)
                            item.PlanFinalText = planText
                        End If
                    ElseIf Not String.IsNullOrWhiteSpace(planText) Then
                        item.PlanStreamText = planText
                    End If

                Case "reasoning"
                    Dim summaryValues = ReadReasoningSummaryValues(itemObject)
                    If summaryValues.Count > 0 Then
                        item.ReasoningSummaryParts.Clear()
                        For Each part In summaryValues
                            item.ReasoningSummaryParts.Add(part)
                        Next
                        item.ReasoningSummaryText = String.Join(Environment.NewLine, summaryValues)
                    Else
                        Dim summaryText = ReadStringFirst(itemObject, "summaryText")
                        If String.IsNullOrWhiteSpace(summaryText) Then
                            summaryText = ReadStringFirst(itemObject, "summary_text")
                        End If

                        If Not String.IsNullOrWhiteSpace(summaryText) Then
                            item.ReasoningSummaryText = summaryText
                        End If
                    End If

                    Dim contentText = ReadReasoningContentText(itemObject)
                    If isFinal Then
                        If Not String.IsNullOrWhiteSpace(contentText) Then
                            completionReplaced = Not StringComparer.Ordinal.Equals(item.ReasoningContentText, contentText)
                            item.ReasoningContentText = contentText
                        End If
                    ElseIf Not String.IsNullOrWhiteSpace(contentText) Then
                        item.ReasoningContentText = contentText
                    End If

                Case "commandExecution"
                    Dim command = ReadStringFirst(itemObject, "command")
                    Dim cwd = ReadStringFirst(itemObject, "cwd")
                    Dim output = ReadStringFirst(itemObject, "aggregatedOutput", "output")

                    If Not String.IsNullOrWhiteSpace(command) Then
                        item.CommandText = command
                    End If

                    If Not String.IsNullOrWhiteSpace(cwd) Then
                        item.CommandCwd = cwd
                    End If

                    If Not String.IsNullOrWhiteSpace(output) Then
                        If isFinal Then
                            completionReplaced = Not StringComparer.Ordinal.Equals(item.CommandOutputText, output)
                        End If
                        item.CommandOutputText = output
                    End If

                    Dim exitCode As Integer
                    If TryReadInteger(itemObject, exitCode, "exitCode", "exit_code") Then
                        item.CommandExitCode = exitCode
                    End If

                    Dim durationMs As Long
                    If TryReadInt64(itemObject, durationMs, "durationMs", "duration_ms", "duration") Then
                        item.CommandDurationMs = durationMs
                    End If

                Case "fileChange"
                    item.FileChangeStatus = NormalizeStatus(ReadStringFirst(itemObject, "status"), item.FileChangeStatus)

                    Dim changes = ReadArray(itemObject, "changes")
                    If changes IsNot Nothing Then
                        item.FileChangeChanges = CloneJsonArray(changes)
                    End If

                    Dim output = ReadStringFirst(itemObject, "aggregatedOutput", "output")
                    If Not String.IsNullOrWhiteSpace(output) Then
                        If isFinal Then
                            completionReplaced = Not StringComparer.Ordinal.Equals(item.FileChangeOutputText, output)
                        End If
                        item.FileChangeOutputText = output
                    End If

                Case Else
                    item.GenericText = BuildGenericItemSummary(itemObject)
            End Select

            Return completionReplaced
        End Function

        Private Sub RefreshItemPendingApprovalState(item As TurnItemRuntimeState)
            If item Is Nothing Then
                Return
            End If

            Dim pendingCount = 0
            For Each approval In _pendingApprovals.Values
                If approval Is Nothing Then
                    Continue For
                End If

                If approval.IsResolved Then
                    Continue For
                End If

                If StringComparer.Ordinal.Equals(approval.ScopedItemKey, item.ScopedItemKey) Then
                    pendingCount += 1
                End If
            Next

            item.PendingApprovalCount = pendingCount
            item.PendingApprovalSummary = If(pendingCount > 0,
                                             $"Approval pending ({pendingCount})",
                                             String.Empty)
        End Sub

        Private Sub ClearPendingApprovalsForTurn(threadId As String, turnId As String)
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            Dim normalizedTurnId = NormalizeIdentifier(turnId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return
            End If

            Dim approvalKeysToRemove As New List(Of String)()

            For Each pair In _pendingApprovals
                Dim approval = pair.Value
                If approval Is Nothing OrElse approval.IsResolved Then
                    Continue For
                End If

                If StringComparer.Ordinal.Equals(approval.ThreadId, normalizedThreadId) AndAlso
                   StringComparer.Ordinal.Equals(approval.TurnId, normalizedTurnId) Then
                    approval.IsResolved = True
                    approval.Decision = "turn_completed"
                    approval.ResolvedAt = DateTimeOffset.UtcNow
                    approvalKeysToRemove.Add(pair.Key)
                End If
            Next

            For Each approvalKey In approvalKeysToRemove
                Dim approval = _pendingApprovals(approvalKey)
                Dim scopedItemKey = approval.ScopedItemKey
                _pendingApprovals.Remove(approvalKey)
                _approvalKeyByRequestId.Remove(approval.RequestIdKey)

                Dim item As TurnItemRuntimeState = Nothing
                If _itemsById.TryGetValue(scopedItemKey, item) Then
                    RefreshItemPendingApprovalState(item)
                End If
            Next
        End Sub

        Private Shared Function BuildTurnKey(threadId As String, turnId As String) As String
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            Dim normalizedTurnId = NormalizeIdentifier(turnId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return String.Empty
            End If

            Return normalizedThreadId & ":" & normalizedTurnId
        End Function

        Private Shared Function BuildItemKey(threadId As String, turnId As String, itemId As String) As String
            Dim turnKey = BuildTurnKey(threadId, turnId)
            Dim normalizedItemId = NormalizeIdentifier(itemId)
            If String.IsNullOrWhiteSpace(turnKey) OrElse String.IsNullOrWhiteSpace(normalizedItemId) Then
                Return String.Empty
            End If

            Return turnKey & ":" & normalizedItemId
        End Function

        Private Shared Function BuildApprovalKey(itemId As String, requestIdKey As String) As String
            Return NormalizeIdentifier(itemId) & "::" & NormalizeIdentifier(requestIdKey)
        End Function

        Private Shared Function BuildRequestIdKey(requestId As JsonNode) As String
            If requestId Is Nothing Then
                Return String.Empty
            End If

            Dim value As String = Nothing
            If TryGetStringValue(requestId, value) AndAlso Not String.IsNullOrWhiteSpace(value) Then
                Return value.Trim()
            End If

            Return requestId.ToJsonString().Trim()
        End Function

        Private Shared Function NormalizeItemType(itemType As String) As String
            Dim normalized = NormalizeIdentifier(itemType)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return String.Empty
            End If

            Select Case normalized.ToLowerInvariant()
                Case "usermessage"
                    Return "userMessage"
                Case "agentmessage"
                    Return "agentMessage"
                Case "plan"
                    Return "plan"
                Case "reasoning"
                    Return "reasoning"
                Case "commandexecution"
                    Return "commandExecution"
                Case "filechange"
                    Return "fileChange"
                Case "mcptoolcall"
                    Return "mcpToolCall"
                Case "collabtoolcall"
                    Return "collabToolCall"
                Case "collabagenttoolcall"
                    Return "collabToolCall"
                Case "websearch"
                    Return "webSearch"
                Case "imageview"
                    Return "imageView"
                Case "enteredreviewmode"
                    Return "enteredReviewMode"
                Case "exitedreviewmode"
                    Return "exitedReviewMode"
                Case "contextcompaction"
                    Return "contextCompaction"
                Case Else
                    Return normalized
            End Select
        End Function

        Private Shared Function ItemTypeFromDeltaKind(deltaKind As ItemDeltaKind) As String
            Select Case deltaKind
                Case ItemDeltaKind.AgentMessage
                    Return "agentMessage"
                Case ItemDeltaKind.Plan
                    Return "plan"
                Case ItemDeltaKind.ReasoningSummaryText,
                     ItemDeltaKind.ReasoningSummaryPartAdded,
                     ItemDeltaKind.ReasoningText
                    Return "reasoning"
                Case ItemDeltaKind.CommandExecutionOutput
                    Return "commandExecution"
                Case ItemDeltaKind.FileChangeOutput
                    Return "fileChange"
                Case Else
                    Return String.Empty
            End Select
        End Function

        Private Shared Function NormalizeStatus(status As String, fallback As String) As String
            Dim normalized = NormalizeIdentifier(status)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return If(fallback, String.Empty)
            End If

            Dim compact = normalized.
                Replace("-", String.Empty, StringComparison.Ordinal).
                Replace("_", String.Empty, StringComparison.Ordinal).
                Replace(" ", String.Empty, StringComparison.Ordinal).
                ToLowerInvariant()

            Select Case compact
                Case "inprogress"
                    Return "in_progress"
                Case "completed"
                    Return "completed"
                Case "failed"
                    Return "failed"
                Case "interrupted"
                    Return "interrupted"
                Case "pending"
                    Return "pending"
                Case Else
                    Return normalized
            End Select
        End Function

        Private Shared Function PreviewDelta(value As String) As String
            Dim text = If(value, String.Empty).Replace(ControlChars.Cr, " "c).
                                          Replace(ControlChars.Lf, " "c).
                                          Replace(ControlChars.Tab, " "c).Trim()
            If text.Length <= 120 Then
                Return text
            End If

            Return text.Substring(0, 117) & "..."
        End Function

        Private Shared Function ReadReasoningContentText(itemObject As JsonObject) As String
            If itemObject Is Nothing Then
                Return String.Empty
            End If

            Dim text = ReadStringFirst(itemObject, "text", "raw_content")
            If Not String.IsNullOrWhiteSpace(text) Then
                Return text
            End If

            Dim contentArray = ReadArray(itemObject, "content")
            If contentArray Is Nothing Then
                contentArray = ReadArray(itemObject, "raw_content")
            End If

            Dim builder As New StringBuilder()
            If contentArray Is Nothing OrElse contentArray.Count = 0 Then
                Return String.Empty
            End If

            For Each contentNode In contentArray
                Dim partText As String = Nothing

                Dim contentObject = TryCast(contentNode, JsonObject)
                If contentObject IsNot Nothing Then
                    partText = ReadStringFirst(contentObject, "text", "delta", "summary")
                ElseIf Not TryGetStringValue(contentNode, partText) Then
                    partText = Nothing
                End If

                If String.IsNullOrWhiteSpace(partText) Then
                    Continue For
                End If

                If builder.Length > 0 Then
                    builder.AppendLine()
                End If

                builder.Append(partText)
            Next

            Return builder.ToString().Trim()
        End Function

        Private Shared Function ReadReasoningSummaryValues(itemObject As JsonObject) As List(Of String)
            Dim parts As New List(Of String)()
            If itemObject Is Nothing Then
                Return parts
            End If

            Dim summaryArrays As JsonArray() = {
                ReadArray(itemObject, "summary"),
                ReadArray(itemObject, "summary_text"),
                ReadArray(itemObject, "summaryParts")
            }

            For Each summaryArray In summaryArrays
                If summaryArray Is Nothing Then
                    Continue For
                End If

                parts.Clear()
                For Each summaryNode In summaryArray
                    Dim partText As String = Nothing
                    Dim summaryObject = TryCast(summaryNode, JsonObject)
                    If summaryObject IsNot Nothing Then
                        partText = ReadStringFirst(summaryObject, "text", "summary")
                    ElseIf Not TryGetStringValue(summaryNode, partText) Then
                        partText = Nothing
                    End If

                    If Not String.IsNullOrWhiteSpace(partText) Then
                        parts.Add(partText)
                    End If
                Next

                If parts.Count > 0 Then
                    Return New List(Of String)(parts)
                End If
            Next

            Return parts
        End Function

        Private Shared Function ReadUserMessageText(itemObject As JsonObject) As String
            If itemObject Is Nothing Then
                Return String.Empty
            End If

            Dim directText = ReadStringFirst(itemObject, "text")
            If Not String.IsNullOrWhiteSpace(directText) Then
                Return directText
            End If

            Dim content = ReadArray(itemObject, "content")
            If content Is Nothing OrElse content.Count = 0 Then
                Return String.Empty
            End If

            Dim parts As New List(Of String)()
            For Each partNode In content
                Dim partObject = TryCast(partNode, JsonObject)
                If partObject Is Nothing Then
                    Continue For
                End If

                Dim partType = ReadStringFirst(partObject, "type")
                Select Case partType
                    Case "text"
                        Dim value = ReadStringFirst(partObject, "text")
                        If Not String.IsNullOrWhiteSpace(value) Then
                            parts.Add(value)
                        End If
                    Case "image"
                        parts.Add($"[image] {ReadStringFirst(partObject, "url")}")
                    Case "localImage"
                        parts.Add($"[localImage] {ReadStringFirst(partObject, "path")}")
                    Case "mention"
                        parts.Add($"[mention] {ReadStringFirst(partObject, "name")}")
                    Case "skill"
                        parts.Add($"[skill] {ReadStringFirst(partObject, "name")}")
                    Case Else
                        Dim value = ReadStringFirst(partObject, "text", "value")
                        If Not String.IsNullOrWhiteSpace(value) Then
                            parts.Add(value)
                        End If
                End Select
            Next

            Return String.Join(Environment.NewLine, parts)
        End Function

        Private Shared Function ReadAgentMessageText(itemObject As JsonObject) As String
            If itemObject Is Nothing Then
                Return String.Empty
            End If

            Dim directText = ReadStringFirst(itemObject, "text")
            If Not String.IsNullOrWhiteSpace(directText) Then
                Return directText
            End If

            Dim content = ReadArray(itemObject, "content")
            If content Is Nothing OrElse content.Count = 0 Then
                Return String.Empty
            End If

            Dim builder As New StringBuilder()
            For Each partNode In content
                Dim partText As String = Nothing
                Dim partObject = TryCast(partNode, JsonObject)
                If partObject IsNot Nothing Then
                    partText = ReadStringFirst(partObject, "text", "value", "delta", "summary")
                ElseIf Not TryGetStringValue(partNode, partText) Then
                    partText = Nothing
                End If

                If String.IsNullOrWhiteSpace(partText) Then
                    Continue For
                End If

                If builder.Length > 0 Then
                    builder.AppendLine()
                End If

                builder.Append(partText)
            Next

            Return builder.ToString().Trim()
        End Function

        Private Shared Function BuildGenericItemSummary(itemObject As JsonObject) As String
            If itemObject Is Nothing Then
                Return String.Empty
            End If

            Dim text = ReadStringFirst(itemObject, "text", "summary", "status")
            If Not String.IsNullOrWhiteSpace(text) Then
                Return text
            End If

            Dim candidateProperties = {"title", "name", "label", "message", "query", "prompt"}
            For Each key In candidateProperties
                text = ReadStringFirst(itemObject, key)
                If Not String.IsNullOrWhiteSpace(text) Then
                    Return text
                End If
            Next

            Return itemObject.ToJsonString()
        End Function

        Private Shared Function ReadAgentMessagePhase(itemObject As JsonObject) As String
            If itemObject Is Nothing Then
                Return String.Empty
            End If

            Return ReadStringFirst(itemObject, "phase")
        End Function

        Private Shared Function ReadStringFirst(obj As JsonObject, ParamArray keys() As String) As String
            If obj Is Nothing OrElse keys Is Nothing Then
                Return String.Empty
            End If

            For Each key In keys
                If String.IsNullOrWhiteSpace(key) Then
                    Continue For
                End If

                Dim node As JsonNode = Nothing
                If obj.TryGetPropertyValue(key, node) AndAlso node IsNot Nothing Then
                    Dim value As String = Nothing
                    If TryGetStringValue(node, value) AndAlso Not String.IsNullOrWhiteSpace(value) Then
                        Return value.Trim()
                    End If
                End If
            Next

            Return String.Empty
        End Function

        Private Shared Function ReadArray(obj As JsonObject, key As String) As JsonArray
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(key, node) Then
                Return Nothing
            End If

            Return TryCast(node, JsonArray)
        End Function

        Private Shared Function ReadNestedStringFirst(obj As JsonObject,
                                                      ParamArray candidatePaths()() As String) As String
            If obj Is Nothing OrElse candidatePaths Is Nothing Then
                Return String.Empty
            End If

            For Each path In candidatePaths
                Dim node = ReadNestedNode(obj, path)
                Dim value As String = Nothing
                If TryGetStringValue(node, value) AndAlso Not String.IsNullOrWhiteSpace(value) Then
                    Return value.Trim()
                End If
            Next

            Return String.Empty
        End Function

        Private Shared Function ReadNestedNode(obj As JsonObject, path As String()) As JsonNode
            If obj Is Nothing OrElse path Is Nothing OrElse path.Length = 0 Then
                Return Nothing
            End If

            Dim current As JsonNode = obj
            For Each segment In path
                Dim currentObject = TryCast(current, JsonObject)
                If currentObject Is Nothing Then
                    Return Nothing
                End If

                If Not currentObject.TryGetPropertyValue(segment, current) Then
                    Return Nothing
                End If
            Next

            Return current
        End Function

        Private Shared Function TryReadInteger(obj As JsonObject, ByRef value As Integer, ParamArray keys() As String) As Boolean
            value = 0
            For Each key In keys
                Dim raw = ReadStringFirst(obj, key)
                If String.IsNullOrWhiteSpace(raw) Then
                    Continue For
                End If

                Dim parsed As Integer
                If Integer.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
                    value = parsed
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function TryReadInt64(obj As JsonObject, ByRef value As Long, ParamArray keys() As String) As Boolean
            value = 0
            For Each key In keys
                Dim raw = ReadStringFirst(obj, key)
                If String.IsNullOrWhiteSpace(raw) Then
                    Continue For
                End If

                Dim parsed As Long
                If Long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
                    value = parsed
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function TryGetStringValue(node As JsonNode, ByRef value As String) As Boolean
            value = String.Empty
            If node Is Nothing Then
                Return False
            End If

            Dim jsonValue = TryCast(node, JsonValue)
            If jsonValue IsNot Nothing Then
                Dim stringValue As String = Nothing
                If jsonValue.TryGetValue(Of String)(stringValue) Then
                    value = If(stringValue, String.Empty)
                    Return True
                End If

                Dim boolValue As Boolean
                If jsonValue.TryGetValue(Of Boolean)(boolValue) Then
                    value = boolValue.ToString(CultureInfo.InvariantCulture)
                    Return True
                End If

                Dim longValue As Long
                If jsonValue.TryGetValue(Of Long)(longValue) Then
                    value = longValue.ToString(CultureInfo.InvariantCulture)
                    Return True
                End If

                value = node.ToJsonString().Trim(""""c)
                Return True
            End If

            value = node.ToString()
            Return Not String.IsNullOrWhiteSpace(value)
        End Function

        Private Shared Function CloneJsonArray(array As JsonArray) As JsonArray
            Return TryCast(CloneJsonNode(array), JsonArray)
        End Function
    End Class
End Namespace
